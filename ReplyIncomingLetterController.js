const Division = require("../../model/Divisions/Divisions");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const TeamLeaders = require("../../model/TeamLeaders/TeamLeaders");
const Directorate = require("../../model/Directorates/Directorates");
const Notification = require("../../model/Notifications/Notification");
const IncomingLetter = require("../../model/IncomingLetters/IncomingLetter");
const ReplyIncomingLetter = require("../../model/ReplyIncomingLetters/ReplyIncomingLetter");
const ForwardIncomingLetter = require("../../model/ForwardIncomingLetters/ForwardIncomingLetter");
const ReplyIncomingLetterHistory = require("../../model/ReplyIncomingLetters/ReplyIncomingLetterHistory");

const { join } = require("path");
const mongoose = require("mongoose");
const formidable = require("formidable");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile } = require("fs/promises");
const {
  appendReplyIncomingLetterPrint,
} = require("../../middleware/replyIncomingLtrPrt");

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

const replyIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CRRPLYINCLTR_API;
    const actualAPIKey = req?.headers?.get_crrplyincltr_api;
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
      const incoming_letter_id = req?.params?.incoming_letter_id;
      const onlineUserList = global?.onlineUserList;

      if (
        !incoming_letter_id ||
        !mongoose.isValidObjectId(incoming_letter_id)
      ) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the incoming letter you want to reply to",
          Message_am: "እባክዎን መልስ መስጠት የሚፈልጉትን ገቢ ደብዳቤ ያቅርቡ",
        });
      }

      const findIncomingLetter = await IncomingLetter.findOne({
        _id: incoming_letter_id,
      });

      if (!findIncomingLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en:
            "The incoming letter that you want to place your reply is not found",
          Message_am: "መልስዎን ማስቀመጥ የሚፈልጉበት ገቢ ደብዳቤ አልተገኘም",
        });
      }

      if (findIncomingLetter?.status === "created") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The incoming letter has not been forwarded, so replies cannot be given.",
          Message_am: "ገቢ ደብዳቤው አልተላከም ስለዚህ ምንም አይነት ምላሽ/reply ማቅረብ አይችሉም።",
        });
      }

      const findForwardIncomingLetter = await ForwardIncomingLetter.findOne({
        incoming_letter_id: incoming_letter_id,
      });

      if (!findForwardIncomingLetter) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Since the incoming letter has not been forwarded yet, replies cannot be sent for this letter",
          Message_am: "ገቢ ደብዳቤው እስካሁን ስላልተላለፈ ለዚህ ደብዳቤ ምላሽ መላክ አይቻልም",
        });
      }

      const findReplyingPersonInForward =
        findForwardIncomingLetter?.path?.filter(
          (item) => item?.to?.toString() === requesterId?.toString()
        );

      if (findReplyingPersonInForward?.length === 0) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You cannot reply to this incoming letter as it has not been forwarded to you",
          Message_am: "ይህ ገቢ ደብዳቤ ለእርስዎ አልተላከም ፤ ስለዚህ ለዚህ ደብዳቤ መልስ መስጠት አይችሉም",
        });
      }

      if (findReplyingPersonInForward?.length > 0) {
        const findNormal = findReplyingPersonInForward?.find(
          (item) => item?.cc === "no"
        );

        if (!findNormal) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You cannot reply to this letter as it was only CC'd to you, not directly forwarded",
            Message_am:
              "ይህንን ደብዳቤ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ መልስ መስጠት አይችሉም",
          });
        }
      }

      const findReplyIncomingLetter = await ReplyIncomingLetter.findOne({
        incoming_letter_id: incoming_letter_id,
      });

      if (findReplyIncomingLetter) {
        const findNormalForward = findReplyingPersonInForward?.find(
          (item) => item?.cc === "no"
        );

        const findSendingUserInReply = findReplyIncomingLetter?.path?.filter(
          (item) => item?.to?.toString() === requesterId?.toString()
        );

        if (findSendingUserInReply?.length > 0) {
          const findNormal = findSendingUserInReply?.find(
            (item) => item?.cc === "no"
          );

          if (!findNormal && !findNormalForward) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en:
                "You cannot reply to this incoming letter as it was only CC'd to you",
              Message_am: "ለእርስዎ CC ብቻ ስለነበር የተደረገው ለዚህ ገቢ ደብዳቤ መልስ መስጠት አይችሉም",
            });
          }
        }
      }

      const form = new formidable.IncomingForm();

      form.parse(req, async (err, fields, files) => {
        if (err) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Error getting data",
            Message_am: "ዳታውን የማግኘት ስህተት",
          });
        }

        const to = fields?.to?.[0];
        const cc = fields?.cc?.[0];
        const remark = fields?.remark?.[0];
        const attachment = files?.attachment?.[0];

        if (!to || !mongoose.isValidObjectId(to)) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Please specify the users to whom you want to send the reply",
            Message_am: "እባክዎ መልስዎን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይምረጡ",
          });
        }

        const findAcceptorUser = await OfficeUser.findOne({ _id: to });

        if (!findAcceptorUser) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: "Recipient user not found",
            Message_am: "ተቀባይ ተጠቃሚ አልተገኘም",
          });
        }

        if (findAcceptorUser?.status !== "active") {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} is currently not active`,
            Message_am: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አክቲቭ ስላልሆኑ ወደ እነርሱ መልስ መላክ አይችሉም`,
          });
        }

        if (findAcceptorUser?._id?.toString() === requesterId?.toString()) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en: "You cannot reply or CC a letter to yourself",
            Message_am: "መልሱን ወደ ራስዎ ማስተላለፍ ወይም CC ማድረግ አይችሉም",
          });
        }

        if (!cc) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Please specify whether you are providing a direct response or just a cc",
            Message_am: "እባክዎ ቀጥተኛ ምላሽ እየሰጡ እንደሆነ ወይም ሲሲ ብቻ ይግለጹ",
          });
        }

        if (cc) {
          if (cc !== "yes" && cc !== "no") {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "Please enter a valid CC type",
              Message_am: "እባክዎ ትክክል የሆነ የCC አይነት ያስገቡ።",
            });
          }
        }

        if (cc === "yes" && attachment) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "CC does not require any attachment",
            Message_am: "CC ምንም አባሪ አይፈልግም",
          });
        }

        if (cc === "no" && !remark && !attachment) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Please provide a remark or attachment to give response",
            Message_am: "ምላሽ ለመስጠት ማብራሪያ ወይም አባሪ ያስፈልጋል",
          });
        }

        const findCCForwardTo = findForwardIncomingLetter?.path?.filter(
          (item) => item?.to?.toString() === to?.toString()
        );

        if (cc === "no" && findCCForwardTo?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: `You cannot reply to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as this incoming letter was not forwarded to him/her`,
            Message_am: `ይህ ገቢ ደብዳቤ ወደ እሱ/ሷ ስላልተላከ ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ምላሽ መስጠት አይችሉም`,
          });
        }

        if (cc === "no" && findCCForwardTo?.length > 0) {
          const findNormal = findCCForwardTo?.find((item) => item?.cc === "no");

          if (!findNormal) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: "You cannot reply to a person who is only CC'd",
              Message_am: "CC ብቻ ለተደረገለት ሰው መልስ መስጠት አይችሉም",
            });
          }
        }

        if (cc === "yes") {
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
                  Message_en: `A division manager is only permitted to CC letters to the directorates, teams, and professionals that are part of their own division. (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                  Message_am: `የዘርፍ ኃላፊ በሱ/ሷ ዘርፍ ውስጥ ካሉት ካልሆነ በስተቀር ለዳይሬክቶሬት፣ ለቡድን ወይም ለባለሙያ ደብዳቤ ምላሽ CC አይችልም። (${
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
                  "A director cannot cc a letter reply to the main director (executive).",
                Message_am: "ዳይሬክተሮች ለዋናው ዳይሬክተር ደብዳቤ ምላሽ CC ማድረግ አይችሉም።",
              });
            }

            if (findAcceptorUser?.level === "DivisionManagers") {
              if (
                findAcceptorUser?.division?.toString() !==
                findRequesterOfficeUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `A director cannot cc a letter reply to a division manager other than his own division manager. (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                  Message_am: `አንድ ዳይሬክተር ከራሱ ዘርፍ ኃላፊ ውጪ ለሌላ ዘርፍ ኃላፊ ደብዳቤ ምላሽ CC ማድረግ አይችልም። (${
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
                  Message_en: `You cannot CC (reply) the letter to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your directorate`,
                  Message_am: `በእርስዎ ዳይሬክቶሬት ውስጥ ስለሌሉ ደብዳቤውን ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} CC (ምላሽ) ማድረግ አይችሉም`,
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
                  "A team leader cannot directly CC a letter reply to the main director or division manager",
                Message_am:
                  "የቡድን መሪ በቀጥታ ወደ ዋና ዳይሬክተር ወይም ዘርፍ ኃላፊ ደብዳቤ ምላሽ CC ማድረግ አይችልም",
              });
            }

            if (findAcceptorUser?.level === "Directors") {
              const findTeamLeadersDirectorate = await Directorate.findOne({
                "members.users": findRequesterOfficeUser?._id,
              });

              if (!findTeamLeadersDirectorate) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}, the team leader, is not found inside any directorate, thus they cannot CC a letter reply to Director ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}፣ የቡድን መሪ፣ በማንኛውም ዳይሬክቶሬት ውስጥ ስለሌለ ደብዳቤውን ለዳይሬክተር ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} CC ምላሽ ማድረግ አይችልም`,
                });
              }

              if (
                findTeamLeadersDirectorate?.manager?.toString() !==
                findAcceptorUser?._id?.toString()
              ) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en:
                    "You are attempting to send a CC reply to a directorate you are not part of",
                  Message_am:
                    "እርስዎ አባል ላልሆኑበት ዳይሬክቶሬት የደብዳቤውን ምላሽ CC ለማድረግ እየሞከሩ ነው።",
                });
              }
            }

            if (findAcceptorUser?.level === "TeamLeaders") {
              if (
                findRequesterOfficeUser?.division?.toString() !==
                findAcceptorUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `A team leader cannot CC a letter reply to another team leader in another division. (${
                    findAcceptorUser?.firstname +
                    " " +
                    findAcceptorUser?.middlename +
                    " " +
                    findAcceptorUser?.lastname
                  })`,
                  Message_am: `የቡድን መሪ በሌላ ዘርፍ ላሉ የቡድን መሪዎች የደብዳቤ ምላሽ CC ማድረግ አይችልም። (${
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
                  Message_en: `You cannot CC the reply of the letter to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your team`,
                  Message_am: `በእርስዎ ቡድን ውስጥ ስለሌሉ የደብዳቤውን ምላሽ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} CC ማድረግ አይችሉም`,
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
                  "A professional cannot directly CC a letter reply to the main director, division manager or director",
                Message_am:
                  "ባለሙያ በቀጥታ ወደ ዋና ዳይሬክተር ፣ ዘርፍ ወይም ዳይሬክተር ኃላፊ ደብዳቤ ምላሽ CC ማድረግ አይችልም",
              });
            }
            if (findAcceptorUser?.level === "TeamLeaders") {
              const findTeamLeaders = await TeamLeaders.findOne({
                "members.users": findRequesterOfficeUser?._id,
              });

              if (!findTeamLeaders) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} cannot send to a team leader that is not his manager. (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                  Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} አስተዳዳሪው ላልሆነ የቡድን መሪ መላክ አይችልም። (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                });
              }

              if (
                findTeamLeaders?.manager?.toString() !==
                findAcceptorUser?._id?.toString()
              ) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} cannot send to a team leader that is not his manager. (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                  Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} አስተዳዳሪው ላልሆነ የቡድን መሪ መላክ አይችልም። (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                });
              }
            }

            if (findAcceptorUser?.level === "Professionals") {
              if (
                findAcceptorUser?.division?.toString() !==
                findRequesterOfficeUser?.division?.toString()
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `You cannot CC a letter reply to a professional who is not part of your division. (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                  Message_am: `የዘርፍዎ አካል ላልሆነ ባለሙያ ደብዳቤ ምላሽ CC ማድረግ አይችሉም። (${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname})`,
                });
              }
            }
          }
        }

        let attachmentName = "";
        if (attachment) {
          if (
            typeof attachment === "object" &&
            (attachment?.mimetype === "application/pdf" ||
              attachment?.mimetype === "application/PDF")
          ) {
            if (attachment?.size > 10 * 1024 * 1024) {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en:
                  "Reply attachment size is too large. Please insert a file less than 10MB.",
                Message_am:
                  "የምላሹ አባሪ መጠኑ በጣም ትልቅ ነው። እባክዎ ከ10ሜባ በታች የሆነ ፋይል ያስገቡ።",
              });
            }
          } else {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Invalid attachment reply format please try again. Only accepts '.pdf'.",
              Message_am: "ልክ ያልሆነ የምላሽ አባሪ እባክዎ እንደገና ይሞክሩ። '.pdf' ብቻ ይቀበላል።",
            });
          }

          const bytes = await readFile(attachment?.filepath);
          const letterReplyAttachmentBuffer = Buffer.from(bytes);
          const uniqueSuffix =
            Date.now() + "-" + Math.round(Math.random() * 1e9);

          const path = join(
            "./",
            "Media",
            "IncomingLetterReplyFiles",
            uniqueSuffix + "-" + attachment?.originalFilename
          );

          attachmentName = uniqueSuffix + "-" + attachment?.originalFilename;

          await writeFile(path, letterReplyAttachmentBuffer);
        }

        if (!findReplyIncomingLetter) {
          const path = [
            {
              replied_date: new Date(),
              from_office_user: requesterId,
              cc: cc,
              to: to,
              attachment: attachmentName,
              remark: remark,
            },
          ];

          const createReplyIncomingLtr = await ReplyIncomingLetter.create({
            incoming_letter_id: incoming_letter_id,
            forward_incoming_letter_id: findForwardIncomingLetter?._id,
            path,
          });

          const updateHistory = [
            {
              updatedByOfficeUser: requesterId,
              action: "create",
            },
          ];

          try {
            await ReplyIncomingLetterHistory.create({
              reply_incoming_letter_id: createReplyIncomingLtr?._id,
              updateHistory,
              history: createReplyIncomingLtr?.toObject(),
            });
          } catch (error) {
            console.log(
              `Reply history for incoming letter with letter number (${findIncomingLetter?.incoming_letter_number}) is not created`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for letter with incoming letter number ${findIncomingLetter?.incoming_letter_number} is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Letter with incoming letter number ${findIncomingLetter?.incoming_letter_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `የደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ባለው ገቢ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
                : `የደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ገቢ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: findAcceptorUser?._id,
            notifcation_type: "IncomingLetter",
            document_id: findIncomingLetter?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(findAcceptorUser?._id, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("incoming_letter_reply_notification", {
              Message_en: `Replies for incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is sent to you.`,
              Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ላለው ደብዳቤ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.CREATED).json({
            Message_en: `Reply for incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
          });
        }

        if (findReplyIncomingLetter) {
          const updateIncomingReply =
            await ReplyIncomingLetter.findOneAndUpdate(
              { _id: findReplyIncomingLetter?._id },
              {
                $push: {
                  path: {
                    replied_date: new Date(),
                    from_office_user: requesterId,
                    cc: cc,
                    to: to,
                    attachment: attachmentName,
                    remark: remark,
                  },
                },
              },
              { new: true }
            );

          if (!updateIncomingReply) {
            return res.status(StatusCodes.EXPECTATION_FAILED).json({
              Message_en: `The attempt to reply to an incoming letter with a letter number ${findIncomingLetter?.incoming_letter_number} for ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} was unsuccessful.`,
              Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}ን ለመመለስ የተደረገው ሙከራ አልተሳካም።`,
            });
          }

          try {
            await ReplyIncomingLetterHistory.findOneAndUpdate(
              { reply_incoming_letter_id: findReplyIncomingLetter?._id },
              {
                $push: {
                  updateHistory: {
                    updatedByOfficeUser: requesterId,
                    action: "update",
                  },
                  history: updateIncomingReply?.toObject(),
                },
              }
            );
          } catch (error) {
            console.log(
              `Reply history of incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is not updated successfully`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ባለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
                : `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: findAcceptorUser?._id,
            notifcation_type: "IncomingLetter",
            document_id: findIncomingLetter?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(findAcceptorUser?._id, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("incoming_letter_reply_notification", {
              Message_en: `Replies for incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is sent to you.`,
              Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ላለው ደብዳቤ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.OK).json({
            Message_en: `Reply for incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
          });
        }
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

const getRepliedIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYINCLTRS_API;
    const actualAPIKey = req?.headers?.get_rplyincltrs_api;
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

      let replyList = [];

      const repliedIncomingLetters = await ReplyIncomingLetter.find({
        "path.to": requesterId,
        "path.cc": "no",
      });

      if (!repliedIncomingLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied incoming letters",
          Message_am: "ምንም መልሶች የተሰጡባቸው ገቢ ደብዳቤዎች የሎዎትም",
        });
      }

      for (const replyPath of repliedIncomingLetters) {
        const pathToRequester = replyPath?.path?.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "no"
        );

        if (pathToRequester) {
          replyList.push(replyPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || 1;
      let sentDate = req?.query?.sent_date || "";
      let receivedDate = req?.query?.received_date || "";
      let incomingLtrNum = req?.query?.incoming_letter_number || "";
      let nimera = req?.query?.nimera || "";
      let attentionFrom = req?.query?.attention_from || "";
      let status = req?.query?.status || "";
      let late = req?.query?.late || "";

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

      const letterIds = replyList?.map(
        (rplyLtr) => rplyLtr?.incoming_letter_id
      );

      const query = {};

      if (attentionFrom) {
        query.attention_from = { $regex: attentionFrom, $options: "i" };
      }
      if (incomingLtrNum) {
        query.incoming_letter_number = {
          $regex: incomingLtrNum,
          $options: "i",
        };
      }
      if (nimera) {
        query.nimera = { $regex: nimera, $options: "i" };
      }
      if (status) {
        query.status = status;
      }
      if (late) {
        query.late = late;
      }
      if (sentDate) {
        query.sent_date = sentDate;
      }
      if (receivedDate) {
        query.received_date = receivedDate;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalIncomingLtrs = await IncomingLetter.countDocuments(query);

      const totalPages = Math.ceil(totalIncomingLtrs / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findIncomingLtrs = await IncomingLetter.find(query)
        .sort({
          createdAt: sortBy,
        })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname",
        });

      if (!findIncomingLtrs) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Incoming letters are not found",
          Message_am: "ገቢ ደብዳቤዎች አልተገኙም",
        });
      }

      const lstOfRepliedIncomingLtr = [];

      for (const items of findIncomingLtrs) {
        const findReplyIncomingLtr = await ReplyIncomingLetter.findOne({
          incoming_letter_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findReplyIncomingLtr) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseReplied: caseFind };

        lstOfRepliedIncomingLtr.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        incomingLetters: lstOfRepliedIncomingLtr,
        totalIncomingLtrs,
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

const getRepliedIncomingLetterCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CCRPLYINCLTRS_API;
    const actualAPIKey = req?.headers?.get_ccrplyincltrs_api;
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

      let replyList = [];

      const repliedIncomingLetters = await ReplyIncomingLetter.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!repliedIncomingLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd incoming letters",
          Message_am: "ምንም CC የተደረጉ ደብዳቤዎች የሎዎትም",
        });
      }

      for (const replyPath of repliedIncomingLetters) {
        const pathToRequester = replyPath?.path?.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "yes"
        );

        if (pathToRequester) {
          replyList.push(replyPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || 1;
      let sentDate = req?.query?.sent_date || "";
      let receivedDate = req?.query?.received_date || "";
      let incomingLtrNum = req?.query?.incoming_letter_number || "";
      let nimera = req?.query?.nimera || "";
      let attentionFrom = req?.query?.attention_from || "";
      let status = req?.query?.status || "";
      let late = req?.query?.late || "";

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

      const letterIds = replyList?.map(
        (rplyLtr) => rplyLtr?.incoming_letter_id
      );

      const query = {};

      if (attentionFrom) {
        query.attention_from = { $regex: attentionFrom, $options: "i" };
      }
      if (incomingLtrNum) {
        query.incoming_letter_number = {
          $regex: incomingLtrNum,
          $options: "i",
        };
      }
      if (nimera) {
        query.nimera = { $regex: nimera, $options: "i" };
      }
      if (status) {
        query.status = status;
      }
      if (late) {
        query.late = late;
      }
      if (sentDate) {
        query.sent_date = sentDate;
      }
      if (receivedDate) {
        query.received_date = receivedDate;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalIncomingLtrs = await IncomingLetter.countDocuments(query);

      const totalPages = Math.ceil(totalIncomingLtrs / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findIncomingLtrs = await IncomingLetter.find(query)
        .sort({
          createdAt: sortBy,
        })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname",
        });

      if (!findIncomingLtrs) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Incoming letters are not found",
          Message_am: "ገቢ ደብዳቤዎች አልተገኙም",
        });
      }

      return res.status(StatusCodes.OK).json({
        incomingLetters: findIncomingLtrs,
        totalIncomingLtrs,
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

const getReplyIncomingLetterPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYINCLTRPATH_API;
    const actualAPIKey = req?.headers?.get_rplyincltrpath_api;
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

      const incoming_letter_id = req?.params?.incoming_letter_id;

      if (
        !incoming_letter_id ||
        !mongoose.isValidObjectId(incoming_letter_id)
      ) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findIncomingLetter = await IncomingLetter.findOne({
        _id: incoming_letter_id,
      });
      const findReplyIncomingLetter = await ReplyIncomingLetter.findOne({
        incoming_letter_id: incoming_letter_id,
      })
        .populate({
          path: "path.from_office_user",
          select: "_id firstname middlename lastname position username level",
        })
        .populate({
          path: "path.to",
          select: "_id firstname middlename lastname position username level",
        });

      if (!findIncomingLetter || !findReplyIncomingLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its replies were not found",
          Message_am: "ደብዳቤው እና ምላሾቹ አልተገኙም",
        });
      }

      const replyLetters = findReplyIncomingLetter?.path;

      return res.status(StatusCodes.OK).json({
        repliedDocs: replyLetters,
        replyId: findReplyIncomingLetter?._id,
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

const printReplyIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYINCLTRPRT_API;
    const actualAPIKey = req?.headers?.get_rplyincltrprt_api;
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

      const id = req?.params?.id;

      if (!id || !mongoose.isValidObjectId(id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findReplyIncomingLetter = await ReplyIncomingLetter.findOne({
        _id: id,
      });
      const findIncomingLetter = await IncomingLetter.findOne({
        _id: findReplyIncomingLetter?.incoming_letter_id,
      });

      if (!findIncomingLetter || !findReplyIncomingLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its replies were not found",
          Message_am: "ደብዳቤው እና ምላሾቹ አልተገኙም",
        });
      }

      const reply_forward_id = req?.params?.reply_forward_id;

      if (!reply_forward_id || !mongoose.isValidObjectId(reply_forward_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the exact reply path you want to print",
          Message_am: "እባክዎ ማተም የሚፈልጉትን ትክክለኛ የምላሽ መንገድ ይግለጹ",
        });
      }

      const replyToPrint = findReplyIncomingLetter?.path?.find(
        (reply) => reply?._id?.toString() === reply_forward_id?.toString()
      );

      if (!replyToPrint) {
        return res.status(StatusCodes.CONFLICT).json({
          Message_en: "The specific reply path does not exist",
          Message_am: "የተፈለገዉ የምላሽ መንገድ የለም",
        });
      }

      const findSenderUser = await OfficeUser.findOne({
        _id: replyToPrint?.from_office_user,
      });

      if (!findSenderUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who replied to this letter (Incoming Letter Number: ${findIncomingLetter?.incoming_letter_number}) is not found among the office administrators.`,
          Message_am: `ለዚህ ደብዳቤ ምላሽ የሰጡት ሰው (የገቢ ደብዳቤ ቁጥር፡ ${findIncomingLetter?.incoming_letter_number}) በቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser.findOne({
        _id: replyToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received the answer to this letter (Letter Number: ${findIncomingLetter?.incoming_letter_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ደብዳቤ መልስ የተቀበለው ሰው (የደብዳቤ ቁጥር፡ ${findIncomingLetter?.incoming_letter_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const sent_date = caseSubDate(replyToPrint?.replied_date);
      const case_num = findIncomingLetter?.incoming_letter_number;
      const sent_from =
        findSenderUser?.firstname +
        " " +
        findSenderUser?.middlename +
        " " +
        findSenderUser?.lastname;
      const sent_to =
        findRecieverUser?.firstname +
        " " +
        findRecieverUser?.middlename +
        " " +
        findRecieverUser?.lastname;
      const cc = replyToPrint?.cc === "yes" ? "አዎ/ነው" : "አይደለም";
      const attachment = replyToPrint?.attachment ? "አባሪ አለው" : "አባሪ የለውም";
      const remark = replyToPrint?.remark;
      const titerImg = findSenderUser?.titer;
      const signatureImg = findSenderUser?.signature;

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
        "ReplyIncomingLetterPrint",
        uniqueSuffix + "-replyincomingletter.pdf"
      );

      const text = {
        sent_date,
        case_num,
        sent_from,
        sent_to,
        attachment,
        cc,
        remark,
        titerImg,
        signatureImg,
      };

      try {
        await appendReplyIncomingLetterPrint(inputPath, text, outputPath);

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
  replyIncomingLetter,
  getRepliedIncomingLetter,
  getRepliedIncomingLetterCC,
  getReplyIncomingLetterPath,
  printReplyIncomingLetter,
};
