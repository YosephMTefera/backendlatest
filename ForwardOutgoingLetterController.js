const Division = require("../../model/Divisions/Divisions");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const TeamLeaders = require("../../model/TeamLeaders/TeamLeaders");
const Directorate = require("../../model/Directorates/Directorates");
const Notification = require("../../model/Notifications/Notification");
const ArchivalUser = require("../../model/ArchivalUsers/ArchivalUsers");
const OutgoingLetter = require("../../model/OutgoingLetters/OutgoingLetter");
const ForwardOutgoingLetter = require("../../model/ForwardOutgoingLetters/ForwardOutgoingLetter");
const ForwardOutgoingLetterHistory = require("../../model/ForwardOutgoingLetters/ForwardOutgoingLetterHistory");

const { join } = require("path");
const cron = require("node-cron");
const mongoose = require("mongoose");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const {
  appendForwardOutgoingLetterPrint,
} = require("../../middleware/forwardOutgoingLtrPrt");

const getUser = (userId, onlineUserList) => {
  return onlineUserList?.find(
    (user) => user?.userId?.toString() === userId?.toString()
  );
};

const caseSubDate = (dateVal) => {
  const newYear = dateVal?.getFullYear();
  const newMonth = dateVal?.getMonth() + 1;
  const newDate = dateVal?.getDate();
  const valDate = ethiopianDate?.toEthiopian(newYear, newMonth, newDate)?.[2];
  const valMonth = ethiopianDate?.toEthiopian(newYear, newMonth, newDate)?.[1];
  const valYear = ethiopianDate?.toEthiopian(newYear, newMonth, newDate)?.[0];
  const ethiopianConvertedDate = `${valDate}/${valMonth}/${valYear}`;
  return ethiopianConvertedDate;
};

const officerOutgoingLetterForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDOFOUTPUTLTR_API;
    const actualAPIKey = req?.headers?.get_frwdofoutputltr_api;
    if (actualAPIKey?.toString() === expectedURLKey?.toString()) {
      const requesterId = req?.user?.id;
      if (!requesterId || !mongoose.isValidObjectId(requesterId)) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const findRequesterOfficeUser = await OfficeUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterOfficeUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterOfficeUser) {
        if (findRequesterOfficeUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      const io = global?.io;
      const outgoing_letter_id = req?.body?.outgoing_letter_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (
        !outgoing_letter_id ||
        !mongoose.isValidObjectId(outgoing_letter_id)
      ) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the outgoing letter to be forwarded",
          Message_am: "እባክዎ የሚላከውን ወጪ ደብዳቤ ያቅርቡ",
        });
      }

      const findOutgoingLtr = await OutgoingLetter.findOne({
        _id: outgoing_letter_id,
      });
      const findForwardOutgoingLtr = await ForwardOutgoingLetter.findOne({
        outgoing_letter_id: outgoing_letter_id,
      });

      if (!findOutgoingLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The outgoing letter to be forwarded/sent is not found",
          Message_am: "የሚተላለፈው/የሚላከው ወጪ ደብዳቤ አልተገኘም",
        });
      }

      if (findOutgoingLtr?.status === "output") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is currently being processed on by the archivals, please wait for the archival to verify the letter.",
          Message_am:
            "ደብዳቤው በአሁኑ ጊዜ በመዝገብ ቤት ፕሮሰስ እየተደረገ ነው፣ እባክዎን ይህን ደብዳቤ ከመላክዎ በፊት መዝገብ ቤቱ ደብዳቤውን ፕሮሰስ እስኪያደርገዉ በትዕግስት ይጠብቁ።",
        });
      }

      const isUserSender = findForwardOutgoingLtr?.path?.some(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (findForwardOutgoingLtr) {
        if (
          !isUserSender &&
          findOutgoingLtr?.createdBy?.toString() !== requesterId?.toString()
        ) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You are unauthorized to forward this letter as it was not initially forwarded to you",
            Message_am: "መጀመሪያ ወደ እርስዎ ስላልተላከ ይህንን ደብዳቤ ማስተላለፍ አልተፈቀደልዎትም",
          });
        }

        const checkIfUserIsCC = findForwardOutgoingLtr?.path?.filter(
          (item) => item?.to?.toString() === requesterId?.toString()
        );

        if (
          findOutgoingLtr?.createdBy?.toString() !== requesterId?.toString()
        ) {
          if (checkIfUserIsCC?.length > 0) {
            const findNormal = checkIfUserIsCC?.find(
              (item) => item?.cc === "no"
            );

            if (!findNormal) {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en:
                  "You cannot forward this outgoing letter as it was only CC'd to you, not directly forwarded",
                Message_am:
                  "ይህንን ወጪ ደብዳቤ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ ማስተላለፍ/መላክ አይችሉም",
              });
            }
          }
        }
      }

      if (
        !findForwardOutgoingLtr &&
        findOutgoingLtr?.createdBy?.toString() !== requesterId?.toString()
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en: `You cannot forward this letter, since you are not the one who created the letter.`,
          Message_am: `ደብዳቤውን የፈጠርከው አንተ ስላልሆንክ ይህን ደብዳቤ ማስተላለፍ አትችልም።`,
        });
      }

      if (!forwardArray || forwardArray?.length === 0) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the users to whom you want to send the letter",
          Message_am: "እባክዎ ደብዳቤውን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይግለጹ",
        });
      }

      if (forwardArray?.length > 0) {
        const forwardToWhomArray = Array.isArray(forwardArray)
          ? forwardArray
          : JSON.parse(forwardArray);

        if (forwardToWhomArray?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Please specify the users to whom you want to send the letter",
            Message_am: "እባክዎ ደብዳቤውን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይግለጹ",
          });
        }

        for (const singlePath of forwardToWhomArray) {
          if (!singlePath?.to || !mongoose.isValidObjectId(singlePath?.to)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please specify the users to whom you want to send the letter",
              Message_am: "እባክዎ ደብዳቤውን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይግለጹ",
            });
          }

          const findAcceptorUser = await OfficeUser.findOne({
            _id: singlePath?.to,
          });

          if (!findAcceptorUser) {
            return res.status(StatusCodes.NOT_FOUND).json({
              Message_en: "Recipient user not found",
              Message_am: "ተቀባይ ተጠቃሚ አልተገኘም",
            });
          }

          if (findAcceptorUser?.status !== "active") {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} is currently not active`,
              Message_am: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አክቲቭ ስላልሆኑ ወደ እነርሱ ደብዳቤ መላክ አይችሉም`,
            });
          }

          if (findAcceptorUser?._id?.toString() === requesterId?.toString()) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: "You can not forward or cc a letter to yourself",
              Message_am: "ደብዳቤን ወደ ራስዎ ማስተላለፍ ወይም CC ማድረግ አይችሉም",
            });
          }

          if (!singlePath?.cc) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please specify if this forward is a normal forward or a cc",
              Message_am: "እባክዎን ይህ ማስተላለፍ/መላክ ሲሲ ነው ወይስ አይደለም የሚለውን ይግለጹ",
            });
          }

          if (singlePath?.cc) {
            if (singlePath?.cc !== "yes" && singlePath?.cc !== "no") {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en: "Please enter a valid cc type",
                Message_am: "እባክዎ ትክክል የሆነ የሲሲ አይነት ያስገቡ።",
              });
            }
          }

          if (!singlePath?.paraph && singlePath?.cc === "no") {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "Please provide a paraph",
              Message_am: "እባክዎን የሚልኩበትን ፓራፍ ያቅርቡ",
            });
          }

          if (singlePath?.paraph && singlePath?.cc === "yes") {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "CC does not require any paraph",
              Message_am: "ሲሲ ምንም ፓራፍ አያስፈልገውም",
            });
          }

          if (singlePath?.paraph) {
            if (!mongoose.isValidObjectId(singlePath?.paraph)) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: "This paraph is not available",
                Message_am: "ይህ ፓራፍ በእርስዎ የፓራፍ ዝርዝር ውስጥ አይገኝም",
              });
            }

            const checkParaphExists = findRequesterOfficeUser?.paraph?.find(
              (p) => p?._id?.toString() === singlePath?.paraph?.toString()
            );

            if (!checkParaphExists) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: "This paraph is not available",
                Message_am: "ይህ ፓራፍ በእርስዎ የፓራፍ ዝርዝር ውስጥ አይገኝም",
              });
            }
          }

          if (findRequesterOfficeUser?.level === "DivisionManagers") {
            if (
              findAcceptorUser?.level === "Directors" ||
              findAcceptorUser?.level === "TeamLeaders" ||
              findAcceptorUser?.level === "Professionals"
            ) {
              if (
                findRequesterOfficeUser?.division?.toString() !==
                findAcceptorUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `A division manager is only permitted to send letters to directorates, teams, and professionals that are part of their own division. (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                  Message_am: `የዘርፍ ኃላፊ በሱ/ሷ ዘርፍ ውስጥ ካሉት ካልሆነ በስተቀር ለዳይሬክቶሬት፣ ለቡድን ወይም ለባለሙያ ደብዳቤ መላክ አይችልም። (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                });
              }
            }
          }

          if (findRequesterOfficeUser?.level === "Directors") {
            const findDirectorate = await Directorate.findOne({
              manager: findRequesterOfficeUser?._id,
            });

            if (!findDirectorate) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: "The director's directorate is not found",
                Message_am: "የዳይሬክተሩ ዳይሬክቶሬት አልተገኘም",
              });
            }

            const findDirectorDivsion = await Division.findOne({
              _id: findRequesterOfficeUser?.division,
            });

            if (!findDirectorDivsion) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: "The director's division is not found",
                Message_am: "የዳይሬክተሩ ዘርፍ አልተገኘም",
              });
            }

            if (
              findAcceptorUser?.level === "MainExecutive" &&
              findDirectorDivsion?.special === "no"
            ) {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en:
                  "A director cannot forward/sent a letter to the main director (executive).",
                Message_am: "ዳይሬክተሮች ለዋናው ዳይሬክተር ደብዳቤ ማስተላለፍ/መላክ አይችሉም።",
              });
            }

            if (findAcceptorUser?.level === "DivisionManagers") {
              if (
                findAcceptorUser?.division?.toString() !==
                findRequesterOfficeUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `A director cannot send a letter to a division manager other than his own division manager. (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                  Message_am: `አንድ ዳይሬክተር ከራሱ ዘርፍ ኃላፊ ውጪ ለሌላ ዘርፍ ኃላፊ ደብዳቤ መላክ/ማስተላለፍ አይችልም። (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                });
              }
            }

            if (
              findAcceptorUser?.level === "TeamLeaders" ||
              findAcceptorUser?.level === "Professionals"
            ) {
              const findMemberInDirectorate = findDirectorate?.members?.find(
                (item) =>
                  item?.users?.toString() === findAcceptorUser?._id?.toString()
              );

              if (!findMemberInDirectorate) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `You cannot forward the letter to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your directorate`,
                  Message_am: `በእርስዎ ዳይሬክቶሬት ውስጥ ስለሌሉ ደብዳቤውን ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ/መላክ አይችሉም`,
                });
              }
            }
          }

          if (findRequesterOfficeUser?.level === "TeamLeaders") {
            const findTeamLeaders = await TeamLeaders.findOne({
              manager: findRequesterOfficeUser?._id,
            });

            if (!findTeamLeaders) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: "The team leader's team is not found",
                Message_am: "የቡድን መሪው ቡድን አልተገኘም",
              });
            }

            if (
              findAcceptorUser?.level === "MainExecutive" ||
              findAcceptorUser?.level === "DivisionManagers"
            ) {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en:
                  "A team leader cannot directly forward to the main director or division manager",
                Message_am:
                  "የቡድን መሪ በቀጥታ ወደ ዋና ዳይሬክተር ወይም ዘርፍ ኃላፊ ማስተላለፍ/መላክ አይችልም",
              });
            }

            if (findAcceptorUser?.level === "Directors") {
              const findTeamLeadersDirectorate = await Directorate.findOne({
                "members.users": findRequesterOfficeUser?._id,
              });

              if (!findTeamLeadersDirectorate) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}, the team leader, is not found inside any directorate, thus they cannot send a letter to Director ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}፣ የቡድን መሪ፣ በማንኛውም ዳይሬክቶሬት ውስጥ ስለሌለ ደብዳቤውን ለዳይሬክተር ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ /መላክ አይችልም`,
                });
              }

              if (
                findTeamLeadersDirectorate?.manager?.toString() !==
                findAcceptorUser?._id?.toString()
              ) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en:
                    "You are attempting to forward to a directorate you are not part of",
                  Message_am: "እርስዎ አባል ላልሆኑበት ዳይሬክቶሬት ለማስተላለፍ እየሞከሩ ነው።",
                });
              }
            }

            if (findAcceptorUser?.level === "TeamLeaders") {
              if (
                findRequesterOfficeUser?.division?.toString() !==
                findAcceptorUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `A team leader cannot forward/send a letter to another team leader in another division. (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                  Message_am: `የቡድን መሪ በሌላ ዘርፍ ላሉ የቡድን መሪዎች ደብዳቤ ማስተላለፍ/መላክ አይችልም። (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                });
              }
            }

            if (findAcceptorUser?.level === "Professionals") {
              const findMemberInTeamLeaders = findTeamLeaders?.members?.find(
                (item) =>
                  item?.users?.toString() === findAcceptorUser?._id?.toString()
              );

              if (!findMemberInTeamLeaders) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `You cannot forward the letter to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your team`,
                  Message_am: `በእርስዎ ቡድን ውስጥ ስለሌሉ ደብዳቤውን ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ/መላክ አይችሉም`,
                });
              }
            }
          }

          if (findRequesterOfficeUser?.level === "Professionals") {
            if (
              findAcceptorUser?.level === "MainExecutive" ||
              findAcceptorUser?.level === "DivisionManagers" ||
              findAcceptorUser?.level === "Directors"
            ) {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en:
                  "A professional cannot directly forward to the main director, division manager and directors",
                Message_am:
                  "ባለሙያዉ በቀጥታ ወደ ዋና ዳይሬክተር፣ ዘርፍ ኃላፊ ወይም ዳይሬክተር ማስተላለፍ/መላክ አይችልም",
              });
            }

            if (findAcceptorUser?.level === "TeamLeaders") {
              const findTeamLeaders = await TeamLeaders.findOne({
                "members.users": findRequesterOfficeUser?._id,
              });

              if (!findTeamLeaders) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} cannot send a letter to a team leader that is not his manager. (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                  Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} አስተዳዳሪው ላልሆነ የቡድን መሪ ደብዳቤ መላክ አይችልም። (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                });
              }

              if (
                findTeamLeaders?.manager?.toString() !==
                findAcceptorUser?._id?.toString()
              ) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} cannot send a letter to a team leader that is not his manager. (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                  Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} አስተዳዳሪው ላልሆነ የቡድን መሪ ደብዳቤ መላክ አይችልም። (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                });
              }
            }

            if (findAcceptorUser?.level === "Professionals") {
              if (
                findAcceptorUser?.division?.toString() !==
                findRequesterOfficeUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `You cannot forward a letter to a professional who is not part or your division. $(${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                  Message_am: `የዘርፍዎ አካል ላልሆነ ባለሙያ ደብዳቤ መላክ አይችሉም። (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                });
              }
            }
          }

          const isUserRecipient = findForwardOutgoingLtr?.path?.filter(
            (item) => item?.to?.toString() === singlePath?.to?.toString()
          );

          if (isUserRecipient?.length > 0) {
            for (const checkRecipient of isUserRecipient) {
              if (checkRecipient?.cc === "yes" && singlePath?.cc === "yes") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This letter has already been cc'd to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ ደብዳቤ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ሲሲ ሆኖ ደርሷዋቸል`,
                });
              }

              if (checkRecipient?.cc === "no" && singlePath?.cc === "no") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This letter has already been forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ ደብዳቤ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ተልኳል እናም ደርሶታል`,
                });
              }
            }
          }
        }

        for (const singlePathSend of forwardToWhomArray) {
          let paraphTitle = "";
          if (singlePathSend?.paraph) {
            const checkParaphExists = findRequesterOfficeUser?.paraph?.find(
              (p) => p?._id?.toString() === singlePathSend?.paraph?.toString()
            );

            paraphTitle = checkParaphExists?.title;
          }

          const findAcceptorUser = await OfficeUser.findOne({
            _id: singlePathSend?.to,
          });

          const findForwardOutgoingLtrCheck =
            await ForwardOutgoingLetter.findOne({
              outgoing_letter_id: outgoing_letter_id,
            });

          if (findForwardOutgoingLtrCheck) {
            const forwardOutgoingLetterToOfficer =
              await ForwardOutgoingLetter.findOneAndUpdate(
                { outgoing_letter_id: outgoing_letter_id },
                {
                  $push: {
                    path: {
                      forwarded_date: new Date(),
                      paraph: paraphTitle,
                      from_office_user: requesterId,
                      cc: singlePathSend?.cc,
                      to: singlePathSend?.to,
                      remark: singlePathSend?.remark,
                    },
                  },
                },
                { new: true }
              );

            if (!forwardOutgoingLetterToOfficer) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The outgoing letter is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
                Message_am: `ወጪ ደብዳቤው ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
              });
            }

            try {
              await ForwardOutgoingLetterHistory.findOneAndUpdate(
                {
                  forward_outgoing_letter_id:
                    forwardOutgoingLetterToOfficer?._id,
                },
                {
                  $push: {
                    updateHistory: {
                      updatedByOfficeUser: requesterId,
                      action: "update",
                    },
                    history: forwardOutgoingLetterToOfficer?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Forward history for outgoing letter with ID ${findOutgoingLtr?._id} is not updated successfully`
              );
            }
          }

          if (!findForwardOutgoingLtrCheck) {
            const path = [
              {
                forwarded_date: new Date(),
                paraph: paraphTitle,
                from_office_user: requesterId,
                cc: singlePathSend?.cc,
                to: singlePathSend?.to,
                remark: singlePathSend?.remark,
              },
            ];

            const newOutgoingLetterForward = await ForwardOutgoingLetter.create(
              {
                outgoing_letter_id: outgoing_letter_id,
                path: path,
              }
            );

            const updateHistory = [
              {
                updatedByOfficeUser: requesterId,
                action: "create",
              },
            ];

            try {
              await ForwardOutgoingLetterHistory.create({
                forward_outgoing_letter_id: newOutgoingLetterForward?._id,
                updateHistory: updateHistory,
                history: newOutgoingLetterForward?.toObject(),
              });
            } catch (error) {
              console.log(
                `Forward history for outgoing letter with ID: ${findOutgoingLtr?._id} is not created successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en:
              singlePathSend?.cc === "no"
                ? `Outgoing letter is forwarded to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Outgoing letter is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              singlePathSend?.cc === "no"
                ? `የወጪ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ ተላልፏል/ተልኳል።`
                : `የወጪ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: findAcceptorUser?._id,
            notifcation_type: "OutgoingLetter",
            document_id: findOutgoingLtr?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(findAcceptorUser?._id, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("outgoing_letter_forward_notification", {
              Message_en: `Outgoing letter is forwarded to you.`,
              Message_am: `ወጪ ደብዳቤው ወደ እርስዎ ተልኳል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `Outgoing letter is forwarded successfully to the recipients.`,
          Message_am: `ወጪ ደብዳቤው ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
        });
      }
    } else {
      return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json();
    }
  } catch (error) {
    return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json({
      Message_en: "Something went wrong please try again",
      Message_am: "ችግር ተፈጥሯል እባክዎ እንደገና ይሞክሩ",
    });
  }
};

const isWeekend = (date) => {
  const day = date.getDay();
  return day === 0 || day === 6;
};

const addWorkingDays = (startDate, days) => {
  let currentDate = new Date(startDate);
  let addedDays = 0;
  while (addedDays < days) {
    currentDate.setDate(currentDate.getDate() + 1);
    if (!isWeekend(currentDate)) {
      addedDays++;
    }
  }
  return currentDate;
};

cron.schedule("30 3 * * *", async () => {
  try {
    const io = global?.io;
    const onlineUserList = global?.onlineUserList;

    const outgoingLetters = await OutgoingLetter.find({
      status: "output",
      late: "no",
    });

    const currentDate = new Date();

    for (const letterItems of outgoingLetters) {
      const registeredDate = letterItems?.output_date;
      const timeLimit = 1;

      const dueDate = addWorkingDays(registeredDate, timeLimit);

      if (currentDate > dueDate) {
        letterItems.late = "yes";
        await letterItems.save();

        const notificationMessage = {
          Message_en: `The outgoing letter with subject ${letterItems?.subject} is submitted to the archivals, but the letter is not verified by the archivals.`,
          Message_am: `የደብዳቤ ርዕስ ${letterItems?.subject} ያለዉ ወጪ ደብዳቤ ወደ መዝገብ ቤት ተልኳል ፤ ነገር ግን በመዝገብ ቤት ኃላፊዎች ወጪ አልተደረገም።`,
        };

        const findMainExe = await OfficeUser.find({
          level: "MainExecutive",
          status: "active",
        });

        const findArchivals = await ArchivalUser.find({
          status: "active",
        });

        if (findMainExe?.length > 0) {
          for (const findExe of findMainExe) {
            await Notification.create({
              office_user: findExe?._id,
              notifcation_type: "LateOutgoingLetter",
              document_id: letterItems?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(findExe?._id, onlineUserList);
            if (user) {
              io.to(user?.socketID).emit("late_outgoing_ltr_notification", {
                Message_en: `${notificationMessage?.Message_en}`,
                Message_am: `${notificationMessage?.Message_am}`,
              });
            }
          }
        }

        if (
          letterItems?.output_by?.toString() !==
          findMainExe?.[0]?._id?.toString()
        ) {
          await Notification.create({
            office_user: letterItems?.output_by,
            notifcation_type: "LateOutgoingLetter",
            document_id: letterItems?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(letterItems?.output_by, onlineUserList);
          if (user) {
            io.to(user?.socketID).emit("late_outgoing_ltr_notification", {
              Message_en: `${notificationMessage?.Message_en}`,
              Message_am: `${notificationMessage?.Message_am}`,
            });
          }
        }

        if (
          letterItems?.createdBy?.toString() !==
            findMainExe?.[0]?._id?.toString() &&
          letterItems?.createdBy?.toString() !==
            letterItems?.output_by?.toString()
        ) {
          await Notification.create({
            office_user: letterItems?.createdBy,
            notifcation_type: "LateOutgoingLetter",
            document_id: letterItems?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(letterItems?.createdBy, onlineUserList);
          if (user) {
            io.to(user?.socketID).emit("late_outgoing_ltr_notification", {
              Message_en: `${notificationMessage?.Message_en}`,
              Message_am: `${notificationMessage?.Message_am}`,
            });
          }
        }

        if (findArchivals?.length > 0) {
          for (const archs of findArchivals) {
            await Notification.create({
              archival_user: archs?._id,
              notifcation_type: "LateOutgoingLetter",
              document_id: letterItems?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(archs?._id, onlineUserList);
            if (user) {
              io.to(user?.socketID).emit("late_outgoing_ltr_notification", {
                Message_en: `${notificationMessage?.Message_en}`,
                Message_am: `${notificationMessage?.Message_am}`,
              });
            }
          }
        }
      }
    }

    console.log("Late outgoing letter checking completed.");
  } catch (error) {
    console.error("Error in late outgoing letter checker: ", error?.message);
  }
});

const getOutgoingForwardLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDGETOUTPUTLTR_API;
    const actualAPIKey = req?.headers?.get_frwdgetoutputltr_api;
    if (actualAPIKey?.toString() === expectedURLKey?.toString()) {
      const requesterId = req?.user?.id;
      if (!requesterId || !mongoose.isValidObjectId(requesterId)) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const findRequesterOfficeUser = await OfficeUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterOfficeUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterOfficeUser) {
        if (findRequesterOfficeUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      let forwardedList = [];

      const forwardedLetters = await ForwardOutgoingLetter.find({
        "path.to": requesterId,
        "path.cc": "no",
      });

      if (!forwardedLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no forwarded letters",
          Message_am: "ምንም የተላለፉ/የተመሩ ደብዳቤዎች የሎዎትም።",
        });
      }

      for (const forwardPath of forwardedLetters) {
        const pathToRequester = forwardPath?.path?.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "no"
        );

        if (pathToRequester) {
          forwardedList?.push(forwardPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || -1;
      let outgoingLtrNum = req?.query?.outgoing_letter_number || "";
      let late = req?.query?.late || "";
      let status = req?.query?.status || "";
      let outputBy = req?.query?.output_by || "";
      let createdBy = req?.query?.createdBy || "";
      let verifiedBy = req?.query?.verified_by || "";
      let outputDate = req?.query?.output_date || "";
      let verifiedDate = req?.query?.verified_date || "";

      if (page <= 0) {
        page = 1;
      }
      if (limit <= 0) {
        limit = 10;
      }
      if (sortBy !== 1 && sortBy !== -1) {
        sortBy = -1;
      }
      if (status === "" || status === null) {
        status = "";
      }
      if (late === "" || late === null) {
        late = "";
      }
      if (!createdBy) {
        createdBy = null;
      }
      if (!outputBy) {
        outputBy = null;
      }
      if (!verifiedBy) {
        verifiedBy = null;
      }
      if (outputBy) {
        if (!outputBy || !mongoose.isValidObjectId(outputBy)) {
          return res
            .status(StatusCodes.NOT_ACCEPTABLE)
            .json({ Message_en: "Invalid request", Message_am: "ልክ ያልሆነ ጥያቄ" });
        }
      }
      if (createdBy) {
        if (!createdBy || !mongoose.isValidObjectId(createdBy)) {
          return res
            .status(StatusCodes.NOT_ACCEPTABLE)
            .json({ Message_en: "Invalid request", Message_am: "ልክ ያልሆነ ጥያቄ" });
        }
      }
      if (verifiedBy) {
        if (!verifiedBy || !mongoose.isValidObjectId(verifiedBy)) {
          return res
            .status(StatusCodes.NOT_ACCEPTABLE)
            .json({ Message_en: "Invalid request", Message_am: "ልክ ያልሆነ ጥያቄ" });
        }
      }

      const findOfficerOutput = await OfficeUser.findOne({ _id: outputBy });
      if (!findOfficerOutput) {
        outputBy = "";
      }

      const findOfficerCreatedBy = await OfficeUser.findOne({ _id: createdBy });
      if (!findOfficerCreatedBy) {
        createdBy = "";
      }

      const findOfficerVerifiedBy = await ArchivalUser.findOne({
        _id: verifiedBy,
      });
      if (!findOfficerVerifiedBy) {
        verifiedBy = "";
      }

      const letterIds = forwardedList?.map(
        (fwdLtr) => fwdLtr?.outgoing_letter_id
      );

      const query = {};

      if (outgoingLtrNum) {
        query.outgoing_letter_number = {
          $regex: outgoingLtrNum,
          $options: "i",
        };
      }
      if (status) {
        query.status = status;
      }
      if (late) {
        query.late = late;
      }
      if (outputDate) {
        query.output_date = outputDate;
      }
      if (verifiedDate) {
        query.verified_date = verifiedDate;
      }
      if (createdBy) {
        query.createdBy = createdBy;
      }
      if (outputBy) {
        query.output_by = outputBy;
      }
      if (verifiedBy) {
        query.verified_by = verifiedBy;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalOutgoingLtrs = await OutgoingLetter.countDocuments(query);

      const totalPages = Math.ceil(totalOutgoingLtrs / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findOutgoingLtrs = await OutgoingLetter.find(query)
        .sort({
          createdAt: sortBy,
        })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "output_by",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "verified_by",
          select: "_id firstname middlename lastname username",
        })
        .populate({
          path: "updated_by.update_officer",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "internal_cc.internal_office",
          select: "_id firstname middlename lastname position",
        });

      if (!findOutgoingLtrs) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Outgoing letters are not found",
          Message_am: "ወጪ ደብዳቤዎች አልተገኙም",
        });
      }

      const lstOfForwardedOutgoingLtr = [];

      for (const items of findOutgoingLtrs) {
        const findFrwdOutgoingLtr = await ForwardOutgoingLetter.findOne({
          outgoing_letter_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findFrwdOutgoingLtr) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseForwarded: caseFind };

        lstOfForwardedOutgoingLtr.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        outgoingLetters: lstOfForwardedOutgoingLtr,
        totalOutgoingLtrs,
        currentPage: page,
        totalPages: totalPages,
      });
    } else {
      return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json();
    }
  } catch (error) {
    return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json({
      Message_en: "Something went wrong please try again",
      Message_am: "ችግር ተፈጥሯል እባክዎ እንደገና ይሞክሩ",
    });
  }
};

const getOutgoingForwardLetterCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDGETCCOUTPUTLTR_API;
    const actualAPIKey = req?.headers?.get_frwdgetccoutputltr_api;
    if (actualAPIKey?.toString() === expectedURLKey?.toString()) {
      const requesterId = req?.user?.id;
      if (!requesterId || !mongoose.isValidObjectId(requesterId)) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const findRequesterOfficeUser = await OfficeUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterOfficeUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterOfficeUser) {
        if (findRequesterOfficeUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      let forwardedList = [];

      const forwardedLetters = await ForwardOutgoingLetter.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!forwardedLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd letters",
          Message_am: "ምንም ግልባጭ የተደረጉ ደብዳቤዎች የሎዎትም",
        });
      }

      for (const forwardPath of forwardedLetters) {
        const pathToRequester = forwardPath?.path?.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "yes"
        );

        if (pathToRequester) {
          forwardedList?.push(forwardPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || -1;
      let outgoingLtrNum = req?.query?.outgoing_letter_number || "";
      let late = req?.query?.late || "";
      let status = req?.query?.status || "";
      let outputBy = req?.query?.output_by || "";
      let createdBy = req?.query?.createdBy || "";
      let verifiedBy = req?.query?.verified_by || "";
      let outputDate = req?.query?.output_date || "";
      let verifiedDate = req?.query?.verified_date || "";

      if (page <= 0) {
        page = 1;
      }
      if (limit <= 0) {
        limit = 10;
      }
      if (sortBy !== 1 && sortBy !== -1) {
        sortBy = -1;
      }
      if (status === "" || status === null) {
        status = "";
      }
      if (late === "" || late === null) {
        late = "";
      }
      if (!createdBy) {
        createdBy = null;
      }
      if (!outputBy) {
        outputBy = null;
      }
      if (!verifiedBy) {
        verifiedBy = null;
      }
      if (outputBy) {
        if (!outputBy || !mongoose.isValidObjectId(outputBy)) {
          return res
            .status(StatusCodes.NOT_ACCEPTABLE)
            .json({ Message_en: "Invalid request", Message_am: "ልክ ያልሆነ ጥያቄ" });
        }
      }
      if (createdBy) {
        if (!createdBy || !mongoose.isValidObjectId(createdBy)) {
          return res
            .status(StatusCodes.NOT_ACCEPTABLE)
            .json({ Message_en: "Invalid request", Message_am: "ልክ ያልሆነ ጥያቄ" });
        }
      }
      if (verifiedBy) {
        if (!verifiedBy || !mongoose.isValidObjectId(verifiedBy)) {
          return res
            .status(StatusCodes.NOT_ACCEPTABLE)
            .json({ Message_en: "Invalid request", Message_am: "ልክ ያልሆነ ጥያቄ" });
        }
      }

      const findOfficerOutput = await OfficeUser.findOne({ _id: outputBy });
      if (!findOfficerOutput) {
        outputBy = "";
      }

      const findOfficerCreatedBy = await OfficeUser.findOne({ _id: createdBy });
      if (!findOfficerCreatedBy) {
        createdBy = "";
      }

      const findOfficerVerifiedBy = await ArchivalUser.findOne({
        _id: verifiedBy,
      });
      if (!findOfficerVerifiedBy) {
        verifiedBy = "";
      }

      const letterIds = forwardedList?.map(
        (fwdLtr) => fwdLtr?.outgoing_letter_id
      );

      const query = {};

      if (outgoingLtrNum) {
        query.outgoing_letter_number = {
          $regex: outgoingLtrNum,
          $options: "i",
        };
      }
      if (status) {
        query.status = status;
      }
      if (late) {
        query.late = late;
      }
      if (outputDate) {
        query.output_date = outputDate;
      }
      if (verifiedDate) {
        query.verified_date = verifiedDate;
      }
      if (createdBy) {
        query.createdBy = createdBy;
      }
      if (outputBy) {
        query.output_by = outputBy;
      }
      if (verifiedBy) {
        query.verified_by = verifiedBy;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalOutgoingLtrs = await OutgoingLetter.countDocuments(query);

      const totalPages = Math.ceil(totalOutgoingLtrs / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findOutgoingLtrs = await OutgoingLetter.find(query)
        .sort({
          createdAt: sortBy,
        })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "output_by",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "verified_by",
          select: "_id firstname middlename lastname username",
        })
        .populate({
          path: "updated_by.update_officer",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "internal_cc.internal_office",
          select: "_id firstname middlename lastname position",
        });

      if (!findOutgoingLtrs) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Outgoing letters are not found",
          Message_am: "ወጪ ደብዳቤዎች አልተገኙም",
        });
      }

      return res.status(StatusCodes.OK).json({
        outgoingLetters: findOutgoingLtrs,
        totalOutgoingLtrs,
        currentPage: page,
        totalPages: totalPages,
      });
    } else {
      return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json();
    }
  } catch (error) {
    return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json({
      Message_en: "Something went wrong please try again",
      Message_am: "ችግር ተፈጥሯል እባክዎ እንደገና ይሞክሩ",
    });
  }
};

const getOutgoingLetterForwardedPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDOUTLTRPATH_API;
    const actualAPIKey = req?.headers?.get_frwdoutltrpath_api;
    if (actualAPIKey?.toString() === expectedURLKey?.toString()) {
      const requesterId = req?.user?.id;
      if (!requesterId || !mongoose.isValidObjectId(requesterId)) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const findRequesterOfficeUser = await OfficeUser.findOne({
        _id: requesterId,
      });

      const findRequesterArchivalUser = await ArchivalUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterOfficeUser && !findRequesterArchivalUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterOfficeUser) {
        if (findRequesterOfficeUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      if (findRequesterArchivalUser) {
        if (findRequesterArchivalUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      const outgoing_letter_id = req?.params?.outgoing_letter_id;

      if (
        !outgoing_letter_id ||
        !mongoose.isValidObjectId(outgoing_letter_id)
      ) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findOutgoingLtrs = await OutgoingLetter.findOne({
        _id: outgoing_letter_id,
      });
      const findForwardOutgoingLtr = await ForwardOutgoingLetter.findOne({
        outgoing_letter_id: outgoing_letter_id,
      })
        .populate({
          path: "path.from_achival_user",
          select: "_id firstname middlename lastname",
        })
        .populate({
          path: "path.from_office_user",
          select: "_id firstname middlename lastname position username level",
        })
        .populate({
          path: "path.to",
          select: "_id firstname middlename lastname position username level",
        });

      if (!findOutgoingLtrs || !findForwardOutgoingLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its forwards were not found",
          Message_am: "ደብዳቤው እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const forwardLetters = findForwardOutgoingLtr?.path;

      return res.status(StatusCodes.OK).json({
        forwardDocs: forwardLetters,
        forwardId: findForwardOutgoingLtr?._id,
      });
    } else {
      return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json();
    }
  } catch (error) {
    return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json({
      Message_en: "Something went wrong please try again",
      Message_am: "ችግር ተፈጥሯል እባክዎ እንደገና ይሞክሩ",
    });
  }
};

const printForwardOutgoingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDOUTPUTLTRPRT_API;
    const actualAPIKey = req?.headers?.get_frwdoutputltrprt_api;
    if (actualAPIKey?.toString() === expectedURLKey?.toString()) {
      const requesterId = req?.user?.id;
      if (!requesterId || !mongoose.isValidObjectId(requesterId)) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const findRequesterOfficeUser = await OfficeUser.findOne({
        _id: requesterId,
      });

      const findRequesterArchivalUser = await ArchivalUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterOfficeUser && !findRequesterArchivalUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterOfficeUser) {
        if (findRequesterOfficeUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      if (findRequesterArchivalUser) {
        if (findRequesterArchivalUser?.status !== "active") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      const id = req?.params?.id;

      if (!id || !mongoose.isValidObjectId(id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findForwardOutgoingLtr = await ForwardOutgoingLetter.findOne({
        _id: id,
      });
      const findOutgoingLtrs = await OutgoingLetter.findOne({
        _id: findForwardOutgoingLtr?.outgoing_letter_id,
      });

      if (!findForwardOutgoingLtr || !findOutgoingLtrs) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its forwards were not found",
          Message_am: "ደብዳቤው እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const print_forward_id = req?.params?.print_forward_id;

      if (!print_forward_id || !mongoose.isValidObjectId(print_forward_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the exact forward path you want to print",
          Message_am: "እባክዎ ማተም የሚፈልጉትን ትክክለኛውን የማስተላለፍ/የመላክ መንገድ ይግለጹ",
        });
      }

      const forwardToPrint = findForwardOutgoingLtr?.path?.find(
        (forward) => forward?._id?.toString() === print_forward_id?.toString()
      );

      if (!forwardToPrint) {
        return res.status(StatusCodes.CONFLICT).json({
          Message_en: "This specific forward path does not exist",
          Message_am: "ይህ የማስተላለፍ/የመላክ መንገድ የለም",
        });
      }

      const findOfficeUserSender = await OfficeUser.findOne({
        _id: forwardToPrint?.from_office_user,
      });

      const findArchivalSender = await ArchivalUser.findOne({
        _id: forwardToPrint?.from_achival_user,
      });

      if (!findOfficeUserSender && !findArchivalSender) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who forwarded this letter (Outgoing Letter Number) is not found among the office administrators or the archival.`,
          Message_am: `ይህንን ደብዳቤ ያስተላለፈው ሰው (የወጪ ደብዳቤ) ከቢሮ ወይም መዝገብ ከቤት አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser?.findOne({
        _id: forwardToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received this letter (Outgoing Letter Number) is not found among the office administrators or the archival.`,
          Message_am: `ይህንን ደብዳቤ የተቀበለው ሰው (የወጪ ደብዳቤ) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      let checkWho = "offUsr";

      if (findArchivalSender) {
        checkWho = "arcUsr";
      }

      const sent_date = caseSubDate(forwardToPrint?.forwarded_date);
      const case_num = findOutgoingLtrs?.outgoing_letter_number
        ? findOutgoingLtrs?.outgoing_letter_number
        : "ቁጥር አልተሰጠዉም";
      let sent_from = "";

      if (findOfficeUserSender) {
        sent_from =
          findOfficeUserSender?.firstname +
          " " +
          findOfficeUserSender?.middlename +
          " " +
          findOfficeUserSender?.lastname;
      }
      if (findArchivalSender) {
        sent_from =
          findArchivalSender?.firstname +
          " " +
          findArchivalSender?.middlename +
          " " +
          findArchivalSender?.lastname;
      }
      const sent_to =
        findRecieverUser?.firstname +
        " " +
        findRecieverUser?.middlename +
        " " +
        findRecieverUser?.lastname;

      const paraph = forwardToPrint?.paraph;
      const cc = forwardToPrint?.cc === "yes" ? "አዎ/ነው" : "አይደለም";
      const remark = forwardToPrint?.remark;

      let titerImg = "";
      if (findOfficeUserSender) {
        titerImg = findOfficeUserSender?.titer;
      }
      if (findArchivalSender) {
        titerImg = findArchivalSender?.titer;
      }

      let signatureImg = "";
      if (findOfficeUserSender) {
        signatureImg = findOfficeUserSender?.signature;
      }
      if (findArchivalSender) {
        signatureImg = findArchivalSender?.signature;
      }

      const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);

      const inputPath = join(
        "./",
        "Media",
        "GenerateFiles",
        "Letterhead_AAHDC.pdf"
      );

      const outputPath = join(
        "./",
        "Media",
        "ForwardOutgoingLetterPrint",
        uniqueSuffix + "-forwardletter.pdf"
      );

      const text = {
        sent_date,
        case_num,
        sent_from,
        sent_to,
        paraph,
        cc,
        remark,
        titerImg,
        signatureImg,
        checkWho,
      };

      try {
        await appendForwardOutgoingLetterPrint(inputPath, text, outputPath);
        const modifiedOutputPath = outputPath
          .replace(/\\/g, "/")
          .replace(/Media\//, "/");

        return res.status(StatusCodes.OK).json(modifiedOutputPath);
      } catch (error) {
        return res.status(StatusCodes.EXPECTATION_FAILED).json({
          Message_en:
            "The system is currently unable to generate the file for printing. Please try again later.",
          Message_am: "ሲስትሙ በአሁኑ ጊዜ ፋይሉን ለህትመት ማመንጨት አልቻለም። እባክዎ ቆየት ብለው ይሞክሩ።",
        });
      }
    } else {
      return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json();
    }
  } catch (error) {
    return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json({
      Message_en: "Something went wrong please try again",
      Message_am: "ችግር ተፈጥሯል እባክዎ እንደገና ይሞክሩ",
    });
  }
};

module.exports = {
  officerOutgoingLetterForward,
  getOutgoingForwardLetter,
  getOutgoingForwardLetterCC,
  getOutgoingLetterForwardedPath,
  printForwardOutgoingLetter,
};
