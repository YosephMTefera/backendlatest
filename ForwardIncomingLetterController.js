const Division = require("../../model/Divisions/Divisions");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const TeamLeaders = require("../../model/TeamLeaders/TeamLeaders");
const Directorate = require("../../model/Directorates/Directorates");
const Notification = require("../../model/Notifications/Notification");
const ArchivalUser = require("../../model/ArchivalUsers/ArchivalUsers");
const IncomingLetter = require("../../model/IncomingLetters/IncomingLetter");
const IncomingLetterHistory = require("../../model/IncomingLetters/IncomingLetterHistory");
const ForwardIncomingLetter = require("../../model/ForwardIncomingLetters/ForwardIncomingLetter");
const ForwardIncomingLetterHistory = require("../../model/ForwardIncomingLetters/ForwardIncomingLetterHistory");

const { join } = require("path");
const cron = require("node-cron");
const mongoose = require("mongoose");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const {
  appendForwardIncLetterPrint,
} = require("../../middleware/forwardIncomingLtrPrt");

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

const archivalIncomingLetterForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_ARCH_INCOM_LTRFRW_API;
    const actualAPIKey = req?.headers?.get_arch_incom_ltrfrw_api;
    if (actualAPIKey?.toString() === expectedURLKey?.toString()) {
      const requesterId = req?.user?.id;
      if (!requesterId || !mongoose.isValidObjectId(requesterId)) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const findRequesterArchivalUser = await ArchivalUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterArchivalUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterArchivalUser?.status !== "active") {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const io = global?.io;
      const incoming_letter_id = req?.body?.incoming_letter_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (
        !incoming_letter_id ||
        !mongoose.isValidObjectId(incoming_letter_id)
      ) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the incoming letter to be forwarded",
          Message_am: "እባክዎ የሚላከውን ገቢ ደብዳቤ ያቅርቡ",
        });
      }

      const findIncomingLetter = await IncomingLetter.findOne({
        _id: incoming_letter_id,
      });
      const findForwardIncomingLetter = await ForwardIncomingLetter.findOne({
        incoming_letter_id: incoming_letter_id,
      });

      if (!findIncomingLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The incoming letter to be forwarded/sent is not found",
          Message_am: "የሚተላለፈው/የሚላከው ገቢ ደብዳቤ አልተገኘም",
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

          if (
            findAcceptorUser?.level !== "MainExecutive" &&
            findAcceptorUser?.level !== "DivisionManagers"
          ) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en:
                "You can only forward the letter to main executive or division managers",
              Message_am: "ደብዳቤውን ለዋና ዳይሬክተር ወይም ለዘርፍ ኃላፊዎች ብቻ ማስተላለፍ ይችላል",
            });
          }

          if (findAcceptorUser?.status !== "active") {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} is currently not active`,
              Message_am: `${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አክቲቭ ስላልሆኑ ወደ እነርሱ ደብዳቤ መላክ አይችሉም`,
            });
          }

          const isUserRecipient = findForwardIncomingLetter?.path?.find(
            (item) => item?.to?.toString() === singlePath?.to?.toString()
          );

          if (isUserRecipient) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: `This letter with letter number ${findIncomingLetter?.incoming_letter_number} has already been forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
              Message_am: `ይህ የደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ አስቀድሞ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል እናም ደርሶታል`,
            });
          }
        }

        // CONTINUE
        if (
          findIncomingLetter?.status === "created" &&
          !findForwardIncomingLetter
        ) {
          let path = [];
          for (const singlePathSend of forwardToWhomArray) {
            const singleAppend = {
              forwarded_date: new Date(),
              from_achival_user: requesterId,
              to: singlePathSend?.to,
              cc: singlePathSend?.cc ? singlePathSend?.cc : "no",
              remark: singlePathSend?.remark,
            };
            path.push(singleAppend);
          }

          const newIncomingLetterForward = await ForwardIncomingLetter.create({
            incoming_letter_id: incoming_letter_id,
            path: path,
          });

          try {
            const updateIncomingForwardLetterHistory = [
              {
                updatedByArchivalUser: requesterId,
                action: "create",
              },
            ];

            await ForwardIncomingLetterHistory.create({
              forward_incoming_letter_id: newIncomingLetterForward?._id,
              updateHistory: updateIncomingForwardLetterHistory,
              history: newIncomingLetterForward?.toObject(),
            });

            const updatedFields = {};

            updatedFields.status = "forwarded";

            const updatedIncomingLetter = await IncomingLetter.findOneAndUpdate(
              { _id: incoming_letter_id },
              updatedFields,
              { new: true }
            );

            await IncomingLetterHistory.findOneAndUpdate(
              { incoming_letter_id: incoming_letter_id },
              {
                $push: {
                  updateHistory: {
                    updatedByArchivalUser: requesterId,
                    action: "update",
                  },
                  history: updatedIncomingLetter?.toObject(),
                },
              }
            );
          } catch (error) {
            await newIncomingLetterForward?.deleteOne({
              incoming_letter_id: incoming_letter_id,
            });

            const findForwardIncomingLetterHistory =
              await ForwardIncomingLetterHistory.findOne({
                forward_incoming_letter_id: newIncomingLetterForward?._id,
              });
            if (findForwardIncomingLetterHistory) {
              await findForwardIncomingLetterHistory?.deleteOne({
                forward_incoming_letter_id: newIncomingLetterForward?._id,
              });
            }

            const findIncomingLetterUpdated = await IncomingLetter.findOne({
              _id: incoming_letter_id,
            });

            if (
              findIncomingLetterUpdated &&
              findIncomingLetterUpdated?.status === "forwarded"
            ) {
              const updatedFields = {};
              updatedFields.status = "created";

              const newUpdatedIncomingLetterSecondV =
                await IncomingLetter.findOneAndUpdate(
                  { _id: incoming_letter_id },
                  updatedFields,
                  { new: true }
                );

              await IncomingLetterHistory.findOneAndUpdate(
                { incoming_letter_id: incoming_letter_id },
                {
                  $push: {
                    updateHistory: {
                      updatedByArchivalUser: requesterId,
                      action: "update",
                    },
                    history: newUpdatedIncomingLetterSecondV?.toObject(),
                  },
                }
              );
            }

            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "The incoming letter was not forwarded successfully. Please try again.",
              Message_am: "ገቢ ደብዳቤው በትክክል አልተላከም። እባክዎን እንደገና ይሞክሩ።",
            });
          }

          const notificationMessage = {
            Message_en: `Incoming Letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded to you from the archival user ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname}`,
            Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ከ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (መዝገብ ቤት) ወደ እርስዎ ተላልፏል/ተልኳል።`,
          };

          for (const singlePathSend of forwardToWhomArray) {
            await Notification.create({
              office_user: singlePathSend?.to,
              notifcation_type: "IncomingLetter",
              document_id: findIncomingLetter?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(singlePathSend?.to, onlineUserList);

            if (user) {
              io.to(user?.socketID).emit(
                "incoming_letter_forward_notification",
                {
                  Message_en: `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded to you.`,
                  Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ወደ እርስዎ ተልኳል።`,
                }
              );
            }
          }

          return res.status(StatusCodes.CREATED).json({
            Message_en: `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded successfully.`,
            Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
          });
        } else if (
          findIncomingLetter?.status === "forwarded" &&
          findForwardIncomingLetter?.path?.length > 0
        ) {
          for (const singlePathSend of forwardToWhomArray) {
            const findAcceptorUser = await OfficeUser.findOne({
              _id: singlePathSend?.to,
            });

            const updatedForwardIncomingLetter =
              await ForwardIncomingLetter.findOneAndUpdate(
                { incoming_letter_id: incoming_letter_id },
                {
                  $push: {
                    path: {
                      forwarded_date: new Date(),
                      from_achival_user: requesterId,
                      to: singlePathSend?.to,
                      cc: singlePathSend?.cc ? singlePathSend?.cc : "no",
                      remark: singlePathSend?.remark,
                    },
                  },
                },
                { new: true }
              );

            if (!updatedForwardIncomingLetter) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The incoming letter is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
                Message_am: `ገቢ ደብዳቤው በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
              });
            }

            try {
              await ForwardIncomingLetterHistory.findOneAndUpdate(
                {
                  forward_incoming_letter_id: updatedForwardIncomingLetter?._id,
                },
                {
                  $push: {
                    updateHistory: {
                      updatedByArchivalUser: requesterId,
                      action: "update",
                    },
                    history: updatedForwardIncomingLetter?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Incoming letter history for letter with letter number (${findIncomingLetter?.incoming_letter_number}) is not updated successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en: `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded to you from the archival user ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname}`,
            Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ከ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (መዝገብ ቤት) ወደ እርስዎ ተላልፏል/ተልኳል።`,
          };

          for (const singlePathSend of forwardToWhomArray) {
            await Notification.create({
              office_user: singlePathSend?.to,
              notifcation_type: "IncomingLetter",
              document_id: findIncomingLetter?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(singlePathSend?.to, onlineUserList);

            if (user) {
              io.to(user?.socketID).emit(
                "incoming_letter_forward_notification",
                {
                  Message_en: `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded to you.`,
                  Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ወደ እርስዎ ተልኳል።`,
                }
              );
            }
          }

          return res.status(StatusCodes.OK).json({
            Message_en: `Internal letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded successfully.`,
            Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
          });
        } else {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "This income letter has a missing forward path or its status has been manipulated. Therefore, it cannot be forwarded.",
            Message_am:
              "ይህ ገቢ ደብዳቤ ያለ አግባብ የተቀየረ ነገር ስላለው ፤ ይህን ደብዳቤ ማስተላለፍ አይቻልም።",
          });
        }
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

cron.schedule("0 3 * * *", async () => {
  try {
    const io = global?.io;
    const onlineUserList = global?.onlineUserList;

    const incomingLetters = await IncomingLetter.find({
      status: "created",
      late: "no",
    });

    const currentDate = new Date();

    for (const letterItems of incomingLetters) {
      const registeredDate = letterItems?.createdAt;
      const timeLimit = 1;

      const dueDate = addWorkingDays(registeredDate, timeLimit);

      if (currentDate > dueDate) {
        letterItems.late = "yes";
        await letterItems.save();

        const notificationMessage = {
          Message_en: `The incoming letter with letter number ${letterItems?.incoming_letter_number} is registered to the system, but the letter is not forwarded to the concerned body on time.`,
          Message_am: `የደብዳቤ ቁጥር ${letterItems?.incoming_letter_number} ያለው ገቢ ደብዳቤ ወደ ስርዓቱ ተመዝግቧል ፤ ነገር ግን ደብዳቤው ከመዝገብ ቤት ወደ ሚመለከተው አካል በጊዜ አልተላለፍም።`,
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
              notifcation_type: "LateIncomingLetter",
              document_id: letterItems?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(findExe?._id, onlineUserList);
            if (user) {
              io.to(user?.socketID).emit("late_incoming_ltr_notification", {
                Message_en: `${notificationMessage?.Message_en}`,
                Message_am: `${notificationMessage?.Message_am}`,
              });
            }
          }
        }

        if (findArchivals?.length > 0) {
          for (const archs of findArchivals) {
            await Notification.create({
              archival_user: archs?._id,
              notifcation_type: "LateIncomingLetter",
              document_id: letterItems?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(archs?._id, onlineUserList);
            if (user) {
              io.to(user?.socketID).emit("late_incoming_ltr_notification", {
                Message_en: `${notificationMessage?.Message_en}`,
                Message_am: `${notificationMessage?.Message_am}`,
              });
            }
          }
        }
      }
    }

    console.log("Late incoming letter checking completed.");
  } catch (error) {
    console.error("Error in late incoming letter checker: ", error?.message);
  }
});

const officerIncomingLetterForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_OFFINCOM_LTRFRW_API;
    const actualAPIKey = req?.headers?.get_offincom_ltrfrw_api;
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
            "Users in this organizational structure (Professionals) cannot forward letters",
          Message_am:
            "በዚህ ድርጅታዊ መዋቅር ውስጥ ያሉ ተጠቃሚዎች (ባለሙያዎች) ደብዳቤዎችን መምራት አይችሉም",
        });
      }

      const io = global?.io;
      const incoming_letter_id = req?.body?.incoming_letter_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (
        !incoming_letter_id ||
        !mongoose.isValidObjectId(incoming_letter_id)
      ) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the incoming letter to be forwarded",
          Message_am: "እባክዎ የሚላከውን ገቢ ደብዳቤ ያቅርቡ",
        });
      }

      const findIncomingLetter = await IncomingLetter.findOne({
        _id: incoming_letter_id,
      });
      const findForwardIncomingLetter = await ForwardIncomingLetter.findOne({
        incoming_letter_id: incoming_letter_id,
      });

      if (!findIncomingLetter || !findForwardIncomingLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The incoming letter to be forwarded/sent is not found",
          Message_am: "የሚተላለፈው/የሚላከው ገቢ ደብዳቤ አልተገኘም",
        });
      }

      if (findIncomingLetter?.status === "created") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "This incoming letter is yet to be forwarded from archivals.",
          Message_am: "ይህ ገቢ ደብዳቤ ከመዝገብ ቤት ገና አልተላለፈም።",
        });
      }

      const isUserSender = findForwardIncomingLetter?.path?.some(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (!isUserSender) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You are unauthorized to forward this letter as it was not initially forwarded to you",
          Message_am: "መጀመሪያ ወደ እርስዎ ስላልተላከ ይህንን ደብዳቤ ማስተላለፍ አልተፈቀደልዎትም",
        });
      }

      const checkIfUserIsCC = findForwardIncomingLetter?.path?.filter(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (checkIfUserIsCC?.length > 0) {
        const findNormal = checkIfUserIsCC?.find((item) => item?.cc === "no");

        if (!findNormal) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You cannot forward this incoming letter as it was only CC'd to you, not directly forwarded",
            Message_am:
              "ይህንን ገቢ ደብዳቤ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ ማስተላለፍ/መላክ አይችሉም",
          });
        }
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
          const isUserRecipient = findForwardIncomingLetter?.path?.filter(
            (item) => item?.to?.toString() === singlePath?.to?.toString()
          );

          if (isUserRecipient?.length > 0) {
            for (const checkRecipient of isUserRecipient) {
              if (checkRecipient?.cc === "yes" && singlePath?.cc === "yes") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This letter with letter number ${findIncomingLetter?.incoming_letter_number} has already been cc'd to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ የደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ሲሲ ሆኖ ደርሷዋቸል`,
                });
              }

              if (checkRecipient?.cc === "no" && singlePath?.cc === "no") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This letter with letter number ${findIncomingLetter?.incoming_letter_number} has already been forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ የደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ተልኳል እናም ደርሶታል`,
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

          const forwardIncomingLetterToOfficer =
            await ForwardIncomingLetter.findOneAndUpdate(
              { incoming_letter_id: incoming_letter_id },
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

          if (!forwardIncomingLetterToOfficer) {
            return res.status(StatusCodes.NOT_FOUND).json({
              Message_en: `The incoming letter is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
              Message_am: `ገቢ ደብዳቤው በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
            });
          }

          try {
            await ForwardIncomingLetterHistory.findOneAndUpdate(
              {
                forward_incoming_letter_id: forwardIncomingLetterToOfficer?._id,
              },
              {
                $push: {
                  updateHistory: {
                    updatedByOfficeUser: requesterId,
                    action: "update",
                  },
                  history: forwardIncomingLetterToOfficer?.toObject(),
                },
              }
            );
          } catch (error) {
            console.log(
              `Forward history for incoming letter with letter number (${findIncomingLetter?.incoming_letter_number}) is not updated successfully`
            );
          }

          const notificationMessage = {
            Message_en:
              singlePathSend?.cc === "no"
                ? `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              singlePathSend?.cc === "no"
                ? `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ ተላልፏል/ተልኳል።`
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
            io.to(user?.socketID).emit("incoming_letter_forward_notification", {
              Message_en: `Letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded to you.`,
              Message_am: `የደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ወደ እርስዎ ተልኳል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `Incoming letter with letter number ${findIncomingLetter?.incoming_letter_number} is forwarded successfully to the recipients.`,
          Message_am: `የገቢ ደብዳቤ ቁጥር ${findIncomingLetter?.incoming_letter_number} ያለው ደብዳቤ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
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

const getIncomingForwardedLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWINC_LTR_API;
    const actualAPIKey = req?.headers?.get_frwinc_ltr_api;
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

      const forwardedLetters = await ForwardIncomingLetter.find({
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

      const letterIds = forwardedList?.map(
        (fwdLtr) => fwdLtr?.incoming_letter_id
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

      const lstOfForwardedIncomingLtrs = [];

      for (const items of findIncomingLtrs) {
        const findFrwdIncLtr = await ForwardIncomingLetter.findOne({
          incoming_letter_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findFrwdIncLtr) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseForwarded: caseFind };

        lstOfForwardedIncomingLtrs.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        incomingLetters: lstOfForwardedIncomingLtrs,
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

const getForwardedIncomingLetterCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_OFFINCOMCC_INCLTR_API;
    const actualAPIKey = req?.headers?.get_offincomcc_incltr_api;
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

      const forwardedLetters = await ForwardIncomingLetter.find({
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
        const pathToRequester = forwardPath.path.find(
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
      let sortBy = parseInt(req?.query?.sort) || 1;
      let sentDate = req?.query?.sent_date || "";
      let receivedDate = req?.query?.received_date || "";
      let incomingLtrNum = req?.query?.incoming_letter_number || "";
      let nimera = req?.query?.nimera || "";
      let attentionFrom = req?.query?.attention_from || "";
      let status = req?.query?.status || "";

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

      const letterIds = forwardedList?.map(
        (fwdLtr) => fwdLtr?.incoming_letter_id
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

const getIncomingLetterForwardedPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDPATH_INCLTR_API;
    const actualAPIKey = req?.headers?.get_frwdpath_incltr_api;
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
      const findForwardIncomingLetter = await ForwardIncomingLetter.findOne({
        incoming_letter_id: incoming_letter_id,
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

      if (!findIncomingLetter || !findForwardIncomingLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its forwards were not found",
          Message_am: "ደብዳቤው እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const forwardLetters = findForwardIncomingLetter?.path;

      return res.status(StatusCodes.OK).json({
        forwardDocs: forwardLetters,
        forwardId: findForwardIncomingLetter?._id,
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

const printForwardIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRTINCFRWD_API;
    const actualAPIKey = req?.headers?.get_prtincfrwd_api;
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

      const findForwardIncomingLetter = await ForwardIncomingLetter.findOne({
        _id: id,
      });
      const findIncomingLetter = await IncomingLetter.findOne({
        _id: findForwardIncomingLetter?.incoming_letter_id,
      });

      if (!findForwardIncomingLetter || !findIncomingLetter) {
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

      const forwardToPrint = findForwardIncomingLetter?.path?.find(
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
          Message_en: `The person who forwarded this letter (Internal Letter Number: ${findIncomingLetter?.incoming_letter_number}) is not found among the office administrators or the archival.`,
          Message_am: `ይህንን ደብዳቤ ያስተላለፈው ሰው (የገቢ ደብዳቤ ቁጥር፡ ${findIncomingLetter?.incoming_letter_number}) ከቢሮ ወይም መዝገብ ከቤት አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser?.findOne({
        _id: forwardToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received this letter (Internal Letter Number: ${findIncomingLetter?.incoming_letter_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ደብዳቤ የተቀበለው ሰው (የገቢ ደብዳቤ ቁጥር፡ ${findIncomingLetter?.incoming_letter_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      let checkWho = "offUsr";

      if (findArchivalSender) {
        checkWho = "arcUsr";
      }

      const sent_date = caseSubDate(forwardToPrint?.forwarded_date);
      const case_num = findIncomingLetter?.incoming_letter_number;
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
        "ForwardIncomingLetterPrint",
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
        await appendForwardIncLetterPrint(inputPath, text, outputPath);
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
  archivalIncomingLetterForward,
  officerIncomingLetterForward,
  getIncomingForwardedLetter,
  getForwardedIncomingLetterCC,
  getIncomingLetterForwardedPath,
  printForwardIncomingLetter,
};
