const Case = require("../../model/Cases/Case");
const CaseList = require("../../model/CaseLists/CaseList");
const Division = require("../../model/Divisions/Divisions");
const CaseHistory = require("../../model/Cases/CaseHistory");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const Directorate = require("../../model/Directorates/Directorates");
const Notification = require("../../model/Notifications/Notification");
const TeamLeaders = require("../../model/TeamLeaders/TeamLeaders");
const ForwardCase = require("../../model/ForwardCases/ForwardCase");
const ForwardCaseHistory = require("../../model/ForwardCases/ForwardCaseHistory");

const { join } = require("path");
const mongoose = require("mongoose");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { appendForwardCasePrint } = require("../../middleware/forwardCasePrint");

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

const officeUserForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CASEFRWD_OFF_API;
    const actualAPIKey = req?.headers?.get_casefrwd_off_api;
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

      if (findRequesterOfficeUser?.level === "Professionals") {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en:
            "Users in this organizational structure (Professionals) cannot forward cases",
          Message_am: "በዚህ ድርጅታዊ መዋቅር ውስጥ ያሉ ተጠቃሚዎች (ባለሙያዎች) ጉዳዮችን መምራት አይችሉም",
        });
      }

      const io = global?.io;
      const case_id = req?.body?.case_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (!case_id || !mongoose.isValidObjectId(case_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the case to be forwarded",
          Message_am: "እባክዎ የሚላከውን ጉዳይ ያቅርቡ",
        });
      }

      const findCase = await Case.findOne({ _id: case_id });
      const findForwardCase = await ForwardCase.findOne({ case_id: case_id });

      if (!findCase || !findForwardCase) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The case to be forwarded is not found",
          Message_am: "የሚላከው ጉዳይ አልተገኘም",
        });
      }

      if (
        findCase?.status === "pending" &&
        findRequesterOfficeUser?.level !== "MainExecutive" &&
        findRequesterOfficeUser?.level !== "DivisionManagers"
      ) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en:
            "In this stage the case can only be forwarded by the main executive or division manager",
          Message_am:
            "በዚህ ደረጃ ጉዳዩን ዋና ዳይሬክተር ወይም ዘርፍ ሃላፊው ብቻ ናቸው ሊመሩት/ሊልኩት የሚችሉት",
        });
      }

      if (
        findCase?.status === "rejected" ||
        findCase?.status === "responded" ||
        findCase?.status === "verified"
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Cases that are rejected or responded/verified, cannot be forwarded",
          Message_am: "ውድቅ የተደረጉ ወይም ምላሽ የተሰጣቸው/የተረጋገጡ ጉዳዮች ሊላኩ አይችሉም",
        });
      }

      if (!forwardArray || forwardArray?.length === 0) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the users to whom you want to send the case",
          Message_am: "እባክዎ ጉዳዩን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይግለጹ",
        });
      }

      const isUserSender = findForwardCase?.path?.some(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (!isUserSender) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You are unauthorized to forward this case as it was not initially forwarded to you",
          Message_am: "መጀመሪያ ወደ እርስዎ ስላልተላከ ይህንን ጉዳይ ማስተላለፍ አልተፈቀደልዎትም",
        });
      }

      const checkIfUserIsCC = findForwardCase?.path?.filter(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (checkIfUserIsCC?.length > 0) {
        const findNormal = checkIfUserIsCC?.find((item) => item?.cc === "no");

        if (!findNormal) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You cannot forward this case as it was only CC'd to you, not directly forwarded",
            Message_am:
              "ይህንን ጉዳይ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ ማስተላለፍ/መላክ አይችሉም",
          });
        }
      }

      if (forwardArray?.length > 0) {
        const forwardToWhomArray = Array.isArray(forwardArray)
          ? forwardArray
          : JSON.parse(forwardArray);

        if (forwardToWhomArray?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Please specify the users to whom you want to send the case",
            Message_am: "እባክዎ ጉዳዩን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይግለጹ",
          });
        }

        for (const singlePath of forwardToWhomArray) {
          if (!singlePath?.to || !mongoose.isValidObjectId(singlePath?.to)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please specify the users to whom you want to send the case",
              Message_am: "እባክዎ ጉዳዩን ለማን መላክ እንደሚፈልጉ ተጠቃሚዎችን ይግለጹ",
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
              Message_am: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አክቲቭ ስላልሆኑ ወደ እነርሱ ጉዳይ መላክ አይችሉም`,
            });
          }

          if (findAcceptorUser?._id?.toString() === requesterId?.toString()) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: "You can not forward or cc a case to yourself",
              Message_am: "ጉዳይን ወደ ራስዎ ማስተላለፍ ወይም CC ማድረግ አይችሉም",
            });
          }

          // Special Division comment
          if (findAcceptorUser?.level !== "MainExecutive") {
            const findUserDivision = await Division.findOne({
              _id: findAcceptorUser?._id,
            });

            if (findUserDivision && findUserDivision?.special === "yes") {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en: `A case cannot be forwarded to officers (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                }) inside a special division`,
                Message_am: `ጉዳዩ በልዩ ዘርፍ ውስጥ ላሉ ኃላፊዎች (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                }) ሊተላለፍ አይችልም`,
              });
            }
          }
          // Special div com

          const isUserRecipient = findForwardCase?.path?.filter(
            (item) => item?.to?.toString() === singlePath?.to?.toString()
          );

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

          if (singlePath?.cc === "no") {
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

              if (
                findAcceptorUser?.level === "TeamLeaders" ||
                findAcceptorUser?.level === "Professionals"
              ) {
                const findMemberInDirectorate = findDirectorate?.members?.find(
                  (item) =>
                    item?.users?.toString() ===
                    findAcceptorUser?._id?.toString()
                );

                if (!findMemberInDirectorate) {
                  return res.status(StatusCodes.FORBIDDEN).json({
                    Message_en: `You cannot forward the case to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your directorate`,
                    Message_am: `በእርስዎ ዳይሬክቶሬት ውስጥ ስለሌሉ ጉዳዩን ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ/መላክ አይችሉም`,
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
                    "A team leader cannot directly forward/cc to the main director or division manager",
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
                    Message_en: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}, the team leader, is not found inside any directorate, thus they cannot send a case to Director ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                    Message_am: `${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}፣ የቡድን መሪ፣ በማንኛውም ዳይሬክቶሬት ውስጥ ስለሌለ ጉዳዩን ለዳይሬክተር ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ /መላክ አይችልም`,
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
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en:
                    "You are not allowed to forward a case to another team leader",
                  Message_am: "ጉዳዩን ወደ ሌላ የቡድን መሪ ማስተላለፍ/መላክ አይፈቀድልዎትም",
                });
              }

              if (findAcceptorUser?.level === "Professionals") {
                const findMemberInTeamLeaders = findTeamLeaders?.members?.find(
                  (item) =>
                    item?.users?.toString() ===
                    findAcceptorUser?._id?.toString()
                );

                if (!findMemberInTeamLeaders) {
                  return res.status(StatusCodes.FORBIDDEN).json({
                    Message_en: `You cannot forward the case to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your team`,
                    Message_am: `በእርስዎ ቡድን ውስጥ ስለሌሉ ጉዳዩን ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ/መላክ አይችሉም`,
                  });
                }
              }
            }
          }

          if (isUserRecipient?.length > 0) {
            for (const checkRecipient of isUserRecipient) {
              if (checkRecipient?.cc === "yes" && singlePath?.cc === "yes") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This case with case number ${findCase?.case_number} has already been cc'd to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ የጉዳይ ቁጥር ${findCase?.case_number} ያለው ጉዳይ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ሲሲ ሆኖ ደርሷዋቸል`,
                });
              }

              if (checkRecipient?.cc === "no" && singlePath?.cc === "no") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This case with case number ${findCase?.case_number} has already been forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ የጉዳይ ቁጥር ${findCase?.case_number} ያለው ጉዳይ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ተልኳል እናም ደርሶታል`,
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

          const forwardCaseToOfficer = await ForwardCase.findOneAndUpdate(
            { _id: findForwardCase?._id },
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

          if (!forwardCaseToOfficer) {
            return res.status(StatusCodes.NOT_FOUND).json({
              Message_en: `The case is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
              Message_am: `ጉዳዩ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
            });
          }

          try {
            await ForwardCaseHistory.findOneAndUpdate(
              { forward_case_id: forwardCaseToOfficer?._id },
              {
                $push: {
                  updateHistory: {
                    updatedByOfficeUser: requesterId,
                    action: "update",
                  },
                  history: forwardCaseToOfficer?.toObject(),
                },
              },
              { new: true }
            );
          } catch (error) {
            console.log(
              `Forward history for case with case number (${findCase?.case_number}) is not updated successfully`
            );
          }

          if (findCase?.status === "pending") {
            if (
              findAcceptorUser?.level === "Directors" ||
              findAcceptorUser?.level === "TeamLeaders" ||
              findAcceptorUser?.level === "Professionals"
            ) {
              const updateCase = await Case.findOneAndUpdate(
                { _id: case_id },
                { status: "ongoing", ongoing_date: new Date() },
                { new: true }
              );

              if (!updateCase) {
                console.log(
                  `Status of case with case number (${findCase?.case_number}) is not found`
                );
              }

              try {
                await CaseHistory.findOneAndUpdate(
                  { case_id: case_id },
                  {
                    $push: {
                      updateHistory: {
                        updatedByOfficeUser: requesterId,
                        action: "update",
                      },
                      history: updateCase?.toObject(),
                    },
                  },
                  { new: true }
                );
              } catch (error) {
                console.log(
                  `Case history of case with case number (${findCase?.case_number}) is not updated successfully`
                );
              }
            }
          }

          const notificationMessage = {
            Message_en:
              singlePathSend?.cc === "no"
                ? `Case with case number ${findCase?.case_number} is forwarded to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Case with case number ${findCase?.case_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              singlePathSend?.cc === "no"
                ? `የጉዳይ ቁጥር ${findCase?.case_number} ያለው ጉዳይ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ ተላልፏል/ተልኳል።`
                : `የጉዳይ ቁጥር ${findCase?.case_number} ያለው ጉዳይ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: findAcceptorUser?._id,
            notifcation_type: "Case",
            document_id: findCase?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(findAcceptorUser?._id, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("case_forward_notification", {
              Message_en: `Case with case number ${findCase?.case_number} is forwarded to you.`,
              Message_am: `የጉዳይ ቁጥር ${findCase?.case_number} ያለው ጉዳይ ወደ እርስዎ ተልኳል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `Case with case number ${findCase?.case_number} is forwarded successfully to the recipients.`,
          Message_am: `የጉዳይ ቁጥር ${findCase?.case_number} ያለው ጉዳይ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
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

const getForwardedCase = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CASEFRWDED_CASE_API;
    const actualAPIKey = req?.headers?.get_casefrwded_case_api;
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

      let forwardedCCList = [];

      const forwardedCases = await ForwardCase.find({
        "path.to": requesterId,
        "path.cc": "no",
      });

      if (!forwardedCases) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no forwarded cases",
          Message_am: "ምንም የተላለፉ/የተመሩ ጉዳዮች የሎዎትም።",
        });
      }

      for (const forwardPath of forwardedCases) {
        const pathToRequester = forwardPath.path.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "no"
        );

        if (pathToRequester) {
          forwardedCCList?.push(forwardPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || 1;
      let status = req?.query?.status || "";
      let late = req?.query?.late || "";
      let division = req?.query?.division || "";
      let caselist = req?.query?.caselist || "";
      let case_number = req?.query?.case_number || "";

      if (page <= 0) {
        page = 1;
      }
      if (limit <= 0) {
        limit = 10;
      }
      if (sortBy !== 1 && sortBy !== -1) {
        sortBy = 1;
      }
      if (status === "" || status === null) {
        status = "";
      }
      if (late === "" || late === null) {
        late = "";
      }
      if (!division) {
        division = null;
      }
      if (!caselist) {
        caselist = null;
      }
      if (division) {
        if (!division || !mongoose.isValidObjectId(division)) {
          return res.status(StatusCodes.NOT_ACCEPTABLE).json({
            Message_en: "Invalid request",
            Message_am: "ልክ ያልሆነ ጥያቄ",
          });
        }
      }

      const findDivision = await Division.findOne({ _id: division });

      if (!findDivision) {
        division = "";
      }

      if (caselist) {
        if (!caselist || !mongoose.isValidObjectId(caselist)) {
          return res.status(StatusCodes.NOT_ACCEPTABLE).json({
            Message_en: "Invalid request",
            Message_am: "ልክ ያልሆነ ጥያቄ",
          });
        }
      }

      const findCaseList = await CaseList.findOne({ _id: caselist });

      if (!findCaseList) {
        caselist = "";
      }

      const caseIds = forwardedCCList?.map((fwdCase) => fwdCase?.case_id);

      const query = {};

      if (status) {
        query.status = status;
      }
      if (late) {
        query.late = late;
      }
      if (division) {
        query.division = division;
      }
      if (caselist) {
        query.caselist = caselist;
      }
      if (case_number) {
        query.case_number = case_number;
      }
      if (caseIds) {
        query._id = { $in: caseIds };
      }

      const totalCases = await Case.countDocuments(query);

      const totalPages = Math.ceil(totalCases / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findCases = await Case.find(query)
        .sort({ createdAt: sortBy })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "customer_id",
          select: "_id firstname middlename lastname",
        })
        .populate({
          path: "window_service_id",
          select: "_id firstname middlename lastname",
        })
        .populate({
          path: "division",
          select: "_id name_en name_am name_or name_sm name_tg name_af",
        })
        .populate({
          path: "caselist",
          select: "_id name_en name_am name_or name_sm name_tg name_af",
        })
        .populate({
          path: "form.list_of_question.question",
          select: "_id name_en name_am name_or name_sm name_tg name_af",
        })
        .populate({
          path: "rejected_by",
          select: "_id firstname middlename lastname username position level",
        })
        .populate({
          path: "responded_by",
          select: "_id firstname middlename lastname username position level",
        })
        .populate({
          path: "schedule_program.schedule.scheduled_by",
          select: "firstname middlename lastname username position level",
        })
        .populate({
          path: "schedule_program.schedule.extended_by",
          select: "firstname middlename lastname username position level",
        })
        .populate({
          path: "verified_by",
          select: "_id firstname middlename lastname username",
        });

      if (!findCases) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no forwarded cases",
          Message_am: "ምንም የተላለፉ/የተመሩ ጉዳዮች የሎዎትም።",
        });
      }

      const lstOfForwardedCases = [];

      for (const items of findCases) {
        const findForwardCase = await ForwardCase.findOne({
          case_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findForwardCase) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseForwarded: caseFind };

        lstOfForwardedCases.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        cases: lstOfForwardedCases,
        totalCases,
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

const getForwardedCaseCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CASEFRWDCC_API;
    const actualAPIKey = req?.headers?.get_casefrwdcc_api;
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

      let forwardedCCList = [];

      const forwardedCases = await ForwardCase.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!forwardedCases) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd cases",
          Message_am: "ምንም የተላለፉ/የተመሩ (CC) ጉዳዮች የሎዎትም።",
        });
      }

      for (const forwardPath of forwardedCases) {
        const pathToRequester = forwardPath.path.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "yes"
        );

        if (pathToRequester) {
          forwardedCCList?.push(forwardPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || 1;
      let status = req?.query?.status || "";
      let late = req?.query?.late || "";
      let division = req?.query?.division || "";
      let caselist = req?.query?.caselist || "";
      let case_number = req?.query?.case_number || "";

      if (page <= 0) {
        page = 1;
      }
      if (limit <= 0) {
        limit = 10;
      }
      if (sortBy !== 1 && sortBy !== -1) {
        sortBy = 1;
      }
      if (status === "" || status === null) {
        status = "";
      }
      if (late === "" || late === null) {
        late = "";
      }
      if (!division) {
        division = null;
      }
      if (!caselist) {
        caselist = null;
      }
      if (division) {
        if (!division || !mongoose.isValidObjectId(division)) {
          return res.status(StatusCodes.NOT_ACCEPTABLE).json({
            Message_en: "Invalid request",
            Message_am: "ልክ ያልሆነ ጥያቄ",
          });
        }
      }

      const findDivision = await Division.findOne({ _id: division });

      if (!findDivision) {
        division = "";
      }

      if (caselist) {
        if (!caselist || !mongoose.isValidObjectId(caselist)) {
          return res.status(StatusCodes.NOT_ACCEPTABLE).json({
            Message_en: "Invalid request",
            Message_am: "ልክ ያልሆነ ጥያቄ",
          });
        }
      }

      const findCaseList = await CaseList.findOne({ _id: caselist });

      if (!findCaseList) {
        caselist = "";
      }

      const caseIds = forwardedCCList?.map((fwdCase) => fwdCase?.case_id);

      const query = {};

      if (status) {
        query.status = status;
      }
      if (late) {
        query.late = late;
      }
      if (division) {
        query.division = division;
      }
      if (caselist) {
        query.caselist = caselist;
      }
      if (case_number) {
        query.case_number = case_number;
      }
      if (caseIds) {
        query._id = { $in: caseIds };
      }

      const totalCases = await Case.countDocuments(query);

      const totalPages = Math.ceil(totalCases / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findCases = await Case.find(query)
        .sort({ createdAt: sortBy })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "customer_id",
          select: "_id firstname middlename lastname",
        })
        .populate({
          path: "window_service_id",
          select: "_id firstname middlename lastname",
        })
        .populate({
          path: "division",
          select: "_id name_en name_am name_or name_sm name_tg name_af",
        })
        .populate({
          path: "caselist",
          select: "_id name_en name_am name_or name_sm name_tg name_af",
        })
        .populate({
          path: "form.list_of_question.question",
          select: "_id name_en name_am name_or name_sm name_tg name_af",
        })
        .populate({
          path: "rejected_by",
          select: "_id firstname middlename lastname username position level",
        })
        .populate({
          path: "responded_by",
          select: "_id firstname middlename lastname username position level",
        })
        .populate({
          path: "schedule_program.schedule.scheduled_by",
          select: "firstname middlename lastname username position level",
        })
        .populate({
          path: "schedule_program.schedule.extended_by",
          select: "firstname middlename lastname username position level",
        })
        .populate({
          path: "verified_by",
          select: "_id firstname middlename lastname username",
        });

      if (!findCases) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd cases",
          Message_am: "ምንም የተላለፉ/የተመሩ (CC) ጉዳዮች የሎዎትም።",
        });
      }

      return res.status(StatusCodes.OK).json({
        cases: findCases,
        totalCases,
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

const getForwardPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CASEFRWDPATH_API;
    const actualAPIKey = req?.headers?.get_casefrwdpath_api;
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

      const case_id = req?.params?.case_id;

      if (!case_id || !mongoose.isValidObjectId(case_id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findCase = await Case.findOne({ _id: case_id });
      const findForwardCase = await ForwardCase.findOne({ case_id: case_id })
        .populate({
          path: "path.from_window_user",
          select: "_id firstname middlename lastname",
        })
        .populate({
          path: "path.from_customer_user",
          select:
            "_id firstname middlename lastname subcity woreda house_number phone gender house_phone_number",
        })
        .populate({
          path: "path.from_office_user",
          select: "_id firstname middlename lastname position username level",
        })
        .populate({
          path: "path.to",
          select: "_id firstname middlename lastname position username level",
        });

      if (!findCase || !findForwardCase) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The case and its forwards were not found",
          Message_am: "ጉዳዩ እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const forwardCases = findForwardCase?.path;

      return res.status(StatusCodes.OK).json({
        forwardDocs: forwardCases,
        forwardId: findForwardCase?._id,
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

const printForwardCase = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRINTCASEFRWD_API;
    const actualAPIKey = req?.headers?.get_printcasefrwd_api;
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

      const findForwardCase = await ForwardCase.findOne({ _id: id });

      const findCase = await Case.findOne({ _id: findForwardCase?.case_id });

      if (!findCase || !findForwardCase) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The case and its forwards were not found",
          Message_am: "ጉዳዩ እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const print_forward_id = req?.params?.print_forward_id;

      if (!print_forward_id || !mongoose.isValidObjectId(print_forward_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the exact forward path you want to print",
          Message_am: "እባክዎ ማተም የሚፈልጉትን ትክክለኛውን የማስተላለፍ/የመላክ መንገድ ይግለጹ",
        });
      }

      const forwardToPrint = findForwardCase?.path?.find(
        (forward) => forward?._id?.toString() === print_forward_id?.toString()
      );

      if (!forwardToPrint) {
        return res.status(StatusCodes.CONFLICT).json({
          Message_en: "This specific forward path does not exist",
          Message_am: "ይህ የማስተላለፍ/የመላክ መንገድ የለም",
        });
      }

      const findSenderUser = await OfficeUser.findOne({
        _id: forwardToPrint?.from_office_user,
      });

      if (!findSenderUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who forwarded this case (Case Number: ${findCase?.case_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ጉዳይ ያስተላለፈው ሰው (የጉዳይ ቁጥር፡ ${findCase?.case_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser?.findOne({
        _id: forwardToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received this case (Case Number: ${findCase?.case_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ጉዳይ የተቀበለው ሰው (የጉዳይ ቁጥር፡ ${findCase?.case_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const sent_date = caseSubDate(forwardToPrint?.forwarded_date);
      const case_num = findCase?.case_number;
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
      const paraph = forwardToPrint?.paraph;
      const cc = forwardToPrint?.cc === "yes" ? "አዎ/ነው" : "አይደለም";
      const remark = forwardToPrint?.remark;
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
        "ForwardCasePrint",
        uniqueSuffix + "-forwardcase.pdf"
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
      };

      try {
        await appendForwardCasePrint(inputPath, text, outputPath);

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
  officeUserForward,
  getForwardedCase,
  getForwardedCaseCC,
  getForwardPath,
  printForwardCase,
};
