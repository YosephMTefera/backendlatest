const Letter = require("../../model/Letters/Letter");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const TeamLeaders = require("../../model/TeamLeaders/TeamLeaders");
const Directorate = require("../../model/Directorates/Directorates");
const Notification = require("../../model/Notifications/Notification");
const LetterHistory = require("../../model/Letters/LetterHistory");
const ForwardLetter = require("../../model/ForwardLetters/ForwardLetter");
const ArchivalUser = require("../../model/ArchivalUsers/ArchivalUsers");
const ForwardLetterHistory = require("../../model/ForwardLetters/ForwardLetterHistory");

const { join } = require("path");
const mongoose = require("mongoose");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const {
  appendForwardLetterPrint,
} = require("../../middleware/forwardLetterPrint");

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

const archivalLetterForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_ARCHLETTERFRWD_API;
    const actualAPIKey = req?.headers?.get_archletterfrwd_api;
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
      const letter_id = req?.body?.letter_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (!letter_id || !mongoose.isValidObjectId(letter_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the letter to be forwarded",
          Message_am: "እባክዎ የሚላከውን ደብዳቤ ያቅርቡ",
        });
      }

      const findLetter = await Letter.findOne({ _id: letter_id });
      const findForwardLetter = await ForwardLetter.findOne({
        letter_id: letter_id,
      });

      if (!findLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter to be forwarded/sent is not found",
          Message_am: "የሚተላለፈው/የሚላከው ደብዳቤ አልተገኘም",
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

          const isUserRecipient = findForwardLetter?.path?.find(
            (item) => item?.to?.toString() === singlePath?.to?.toString()
          );

          if (isUserRecipient) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: `This letter with letter number ${findLetter?.letter_number} has already been forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
              Message_am: `ይህ የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ አስቀድሞ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል እናም ደርሶታል`,
            });
          }
        }

        if (findLetter?.status === "created" && !findForwardLetter) {
          let path = [];
          for (const singlePathSend of forwardToWhomArray) {
            const singleAppend = {
              forwarded_date: new Date(),
              from_achival_user: requesterId,
              to: singlePathSend?.to,
              cc: singlePathSend?.cc === "yes" ? singlePathSend?.cc : "no",
              remark: singlePathSend?.remark,
            };
            path.push(singleAppend);
          }

          const newLetterForward = await ForwardLetter.create({
            letter_id: letter_id,
            path: path,
          });

          try {
            const updateForwardLetterHistory = [
              {
                updatedByArchivalUser: requesterId,
                action: "create",
              },
            ];

            await ForwardLetterHistory.create({
              forward_letter_id: newLetterForward?._id,
              updateHistory: updateForwardLetterHistory,
              history: newLetterForward?.toObject(),
            });

            const updatedFields = {};

            updatedFields.status = "forwarded";

            const updatedLetter = await Letter.findOneAndUpdate(
              { _id: letter_id },
              updatedFields,
              { new: true }
            );

            await LetterHistory.findOneAndUpdate(
              { letter_id: letter_id },
              {
                $push: {
                  updateHistory: {
                    updatedByArchivalUser: requesterId,
                    action: "update",
                  },
                  history: updatedLetter?.toObject(),
                },
              }
            );
          } catch (error) {
            await newLetterForward?.deleteOne({ letter_id: letter_id });

            const findForwardLetterHistory = await ForwardLetterHistory.findOne(
              { forward_letter_id: newLetterForward?._id }
            );

            if (findForwardLetterHistory) {
              await findForwardLetterHistory?.deleteOne({
                forward_letter_id: newLetterForward?._id,
              });
            }

            const findLetterUpdated = await Letter.findOne({ _id: letter_id });

            if (
              findLetterUpdated &&
              findLetterUpdated?.status === "forwarded"
            ) {
              const updatedFields = {};
              updatedFields.status = "created";
              const newUpdatedLetterSecondV = await Letter.findOneAndUpdate(
                { _id: letter_id },
                updatedFields,
                { new: true }
              );

              await LetterHistory.findOneAndUpdate(
                { letter_id: letter_id },
                {
                  $push: {
                    updateHistory: {
                      updatedByArchivalUser: requesterId,
                      action: "update",
                    },
                    history: newUpdatedLetterSecondV?.toObject(),
                  },
                }
              );
            }

            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "The letter was not forwarded successfully. Please try again.",
              Message_am: "ደብዳቤው በትክክል አልተላከም። እባክዎን እንደገና ይሞክሩ።",
            });
          }

          const notificationMessage = {
            Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded to you from the archival user ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname}`,
            Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ከ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (መዝገብ ቤት) ወደ እርስዎ ተላልፏል/ተልኳል።`,
          };

          for (const singlePathSend of forwardToWhomArray) {
            await Notification.create({
              office_user: singlePathSend?.to,
              notifcation_type: "Letter",
              document_id: findLetter?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(singlePathSend?.to, onlineUserList);

            if (user) {
              io.to(user?.socketID).emit("letter_forward_notification", {
                Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded to you.`,
                Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ወደ እርስዎ ተልኳል።`,
              });
            }
          }

          return res.status(StatusCodes.CREATED).json({
            Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded successfully.`,
            Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
          });
        } else if (
          findLetter?.status === "forwarded" &&
          findForwardLetter?.path?.length > 0
        ) {
          for (const singlePathSend of forwardToWhomArray) {
            const findAcceptorUser = await OfficeUser.findOne({
              _id: singlePathSend?.to,
            });

            const updatedForwardLetter = await ForwardLetter.findOneAndUpdate(
              { letter_id: letter_id },
              {
                $push: {
                  path: {
                    forwarded_date: new Date(),
                    from_achival_user: requesterId,
                    to: singlePathSend?.to,
                    cc:
                      singlePathSend?.cc === "yes" ? singlePathSend?.cc : "no",
                    remark: singlePathSend?.remark,
                  },
                },
              },
              { new: true }
            );

            if (!updatedForwardLetter) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The letter is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
                Message_am: `ደብዳቤው በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
              });
            }

            try {
              await ForwardLetterHistory.findOneAndUpdate(
                { forward_letter_id: updatedForwardLetter?._id },
                {
                  $push: {
                    updateHistory: {
                      updatedByArchivalUser: requesterId,
                      action: "update",
                    },
                    history: updatedForwardLetter?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Letter history for letter with letter number (${findLetter?.letter_number}) is not updated successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded to you from the archival user ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname}`,
            Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ከ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (መዝገብ ቤት) ወደ እርስዎ ተላልፏል/ተልኳል።`,
          };

          for (const singlePathSend of forwardToWhomArray) {
            await Notification.create({
              office_user: singlePathSend?.to,
              notifcation_type: "Letter",
              document_id: findLetter?._id,
              message_en: notificationMessage?.Message_en,
              message_am: notificationMessage?.Message_am,
            });

            const user = getUser(singlePathSend?.to, onlineUserList);

            if (user) {
              io.to(user?.socketID).emit("letter_forward_notification", {
                Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded to you.`,
                Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ወደ እርስዎ ተልኳል።`,
              });
            }
          }

          return res.status(StatusCodes.OK).json({
            Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded successfully.`,
            Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
          });
        } else {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "This letter has a missing forward path or its status has been manipulated. Therefore, it cannot be forwarded.",
            Message_am:
              "ይህ ደብዳቤ ያለ አግባብ የተቀየረ ነገር ስላለው ፤ ይህን ደብዳቤ ማስተላለፍ አይቻልም።",
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

const officerLetterForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_OFFLETTERFRWD_API;
    const actualAPIKey = req?.headers?.get_offletterfrwd_api;
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

      // if (findRequesterOfficeUser?.level === "Professionals") {
      //   return res.status(StatusCodes.UNAUTHORIZED).json({
      //     Message_en:
      //       "Users in this organizational structure (Professionals) cannot forward letters",
      //     Message_am:
      //       "በዚህ ድርጅታዊ መዋቅር ውስጥ ያሉ ተጠቃሚዎች (ባለሙያዎች) ደብዳቤዎችን መምራት አይችሉም",
      //   });
      // }

      const io = global?.io;
      const letter_id = req?.body?.letter_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (!letter_id || !mongoose.isValidObjectId(letter_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the letter to be forwarded",
          Message_am: "እባክዎ የሚላከውን ደብዳቤ ያቅርቡ",
        });
      }

      const findLetter = await Letter.findOne({ _id: letter_id });
      const findForwardLetter = await ForwardLetter.findOne({
        letter_id: letter_id,
      });

      if (!findLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter to be forwarded/sent is not found",
          Message_am: "የሚተላለፈው/የሚላከው ደብዳቤ አልተገኘም",
        });
      }

      const isUserSender = findForwardLetter?.path?.some(
        (item) =>
          item?.to?.toString() === requesterId?.toString() ||
          item?.from_office_user?.toString() === requesterId?.toString()
      );

      if (
        !findForwardLetter &&
        findRequesterOfficeUser?.level !== "MainExecutive" &&
        findRequesterOfficeUser?.level !== "DivisionManagers"
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You are unauthorized to forward this letter as it was not initially forwarded to you",
          Message_am: "መጀመሪያ ወደ እርስዎ ስላልተላከ ይህንን ደብዳቤ ማስተላለፍ አልተፈቀደልዎትም",
        });
      }

      if (findForwardLetter && !isUserSender) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You are unauthorized to forward this letter as it was not initially forwarded to you",
          Message_am: "መጀመሪያ ወደ እርስዎ ስላልተላከ ይህንን ደብዳቤ ማስተላለፍ አልተፈቀደልዎትም",
        });
      }

      const checkIfUserIsCC = findForwardLetter?.path?.filter(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (
        checkIfUserIsCC?.length > 0 &&
        findRequesterOfficeUser?.level !== "MainExecutive" &&
        findRequesterOfficeUser?.level !== "DivisionManagers"
      ) {
        const findNormal = checkIfUserIsCC?.find((item) => item?.cc === "no");

        if (!findNormal) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You cannot forward this letter as it was only CC'd to you, not directly forwarded",
            Message_am:
              "ይህንን ደብዳቤ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ ማስተላለፍ/መላክ አይችሉም",
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
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en:
                    "You are not allowed to forward a letter to another team leader",
                  Message_am: "ደብዳቤውን ወደ ሌላ የቡድን መሪ ማስተላለፍ/መላክ አይፈቀድልዎትም",
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
                    Message_en: `You cannot forward the letter to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as they are not in your team`,
                    Message_am: `በእርስዎ ቡድን ውስጥ ስለሌሉ ደብዳቤውን ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ማስተላለፍ/መላክ አይችሉም`,
                  });
                }
              }
            }

            if (findRequesterOfficeUser?.level === "Professionals") {
              const findTeamLeaders = await TeamLeaders.findOne({
                "members.users": findRequesterOfficeUser?._id,
              });

              if (!findTeamLeaders) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en: "The professional's team is not found",
                  Message_am: "የባለሙያዉ ቡድን አልተገኘም",
                });
              }

              if (
                findAcceptorUser?.level === "MainExecutive" ||
                findAcceptorUser?.level === "DivisionManagers" ||
                findAcceptorUser?.level === "Directors"
              ) {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en:
                    "A professional cannot directly forward to the main director, division managers, and directors",
                  Message_am:
                    "ባለሙያ በቀጥታ ወደ ዋና ዳይሬክተር ፣ ዘርፍ ኃላፊ ወይም ወደ ዳይሬክተር ማስተላለፍ/መላክ አይችልም",
                });
              }

              if (findAcceptorUser?.level === "TeamLeaders") {
                if (
                  findTeamLeaders?.manager?.toString() ===
                  findAcceptorUser?._id?.toString()
                ) {
                  return res.status(StatusCodes.FORBIDDEN).json({
                    Message_en:
                      "A professional cannot forward letters to team leaders that they are not part of",
                    Message_am:
                      "ባለሙያዎች አባል ባልሆኑበት ቡድን ዉስጥ ላለ ቡድን መሪ ደብዳቤ ማስተላለፍ/መላክ አይችልም",
                  });
                }
              }

              if (findAcceptorUser?.level === "Professionals") {
                if (
                  findAcceptorUser?.division?.toString() !==
                  findRequesterOfficeUser?.division?.toString()
                ) {
                  return res.status(StatusCodes.FORBIDDEN).json({
                    Message_en:
                      "A professional cannot forward letters to professionals in another division",
                    Message_am:
                      "ባለሙያዎች በሌላ ዘርፍ ዉስጥ ላሉ ባለሙያዎች ደብዳቤ ማስተላለፍ/መላክ አይችልም",
                  });
                }
              }
            }
          }

          const isUserRecipient = findForwardLetter?.path?.filter(
            (item) => item?.to?.toString() === singlePath?.to?.toString()
          );

          if (isUserRecipient?.length > 0) {
            for (const checkRecipient of isUserRecipient) {
              if (checkRecipient?.cc === "yes" && singlePath?.cc === "yes") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This letter with letter number ${findLetter?.letter_number} has already been cc'd to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ሲሲ ሆኖ ደርሷዋቸል`,
                });
              }

              if (checkRecipient?.cc === "no" && singlePath?.cc === "no") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `This letter with letter number ${findLetter?.letter_number} has already been forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}`,
                  Message_am: `ይህ የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ለ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አስቀድሞ ተልኳል እናም ደርሶታል`,
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

          const findForwardLetterCheck = await ForwardLetter.findOne({
            letter_id: letter_id,
          });

          if (!findForwardLetterCheck) {
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

            const newLetterForward = await ForwardLetter.create({
              letter_id: letter_id,
              path: path,
            });

            const updateHistory = [
              {
                updatedByOfficeUser: requesterId,
                action: "create",
              },
            ];

            try {
              await ForwardLetterHistory.create({
                forward_letter_id: newLetterForward?._id,
                updateHistory,
                history: newLetterForward?.toObject(),
              });

              const updatedFields = {};

              updatedFields.status = "forwarded";

              const updatedLetter = await Letter.findOneAndUpdate(
                { _id: letter_id },
                updatedFields,
                { new: true }
              );

              await LetterHistory.findOneAndUpdate(
                { letter_id: letter_id },
                {
                  $push: {
                    updateHistory: {
                      updatedByOfficeUser: requesterId,
                      action: "update",
                    },
                    history: updatedLetter?.toObject(),
                  },
                }
              );
            } catch (error) {
              `Forward history for letter with letter number (${findLetter?.letter_number}) is not created successfully`;
            }
          }

          if (findForwardLetterCheck) {
            const forwardLetterToOfficer = await ForwardLetter.findOneAndUpdate(
              { letter_id: letter_id },
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

            if (!forwardLetterToOfficer) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The letter is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
                Message_am: `ደብዳቤው በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
              });
            }

            try {
              await ForwardLetterHistory.findOneAndUpdate(
                { forward_letter_id: forwardLetterToOfficer?._id },
                {
                  $push: {
                    updateHistory: {
                      updatedByOfficeUser: requesterId,
                      action: "update",
                    },
                    history: forwardLetterToOfficer?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Forward history for letter with letter number (${findLetter?.letter_number}) is not updated successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en:
              singlePathSend?.cc === "no"
                ? `Letter with letter number ${findLetter?.letter_number} is forwarded to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Letter with letter number ${findLetter?.letter_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              singlePathSend?.cc === "no"
                ? `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ ተላልፏል/ተልኳል።`
                : `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: findAcceptorUser?._id,
            notifcation_type: "Letter",
            document_id: findLetter?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(findAcceptorUser?._id, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("letter_forward_notification", {
              Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded to you.`,
              Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ወደ እርስዎ ተልኳል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `Letter with letter number ${findLetter?.letter_number} is forwarded successfully to the recipients.`,
          Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ያለው ደብዳቤ ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።`,
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

const getForwardedLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDLETTERS_API;
    const actualAPIKey = req?.headers?.get_frwdletters_api;
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

      const forwardedLetters = await ForwardLetter.find({
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
        const pathToRequester = forwardPath.path.find(
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
      let letterNum = req?.query?.letter_number || "";
      let nimera = req?.query?.nimera || "";
      let letterType = req?.query?.letter_type || "";
      let sentFrom = req?.query?.sent_from || "";
      let sentTo = req?.query?.sent_to || "";
      let sentDate = req?.query?.letter_sent_date || "";
      let status = req?.query?.status || "";

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
      if (letterType === "" || letterType === null) {
        letterType = "";
      }

      const letterIds = forwardedList?.map((fwdLetter) => fwdLetter?.letter_id);

      const query = {};

      if (letterNum) {
        query.letter_number = { $regex: letterNum, $options: "i" };
      }
      if (nimera) {
        query.nimera = { $regex: nimera, $options: "i" };
      }
      if (letterType) {
        query.letter_type = letterType;
      }
      if (sentFrom) {
        query.sent_from = { $regex: sentFrom, $options: "i" };
      }
      if (sentTo) {
        query.sent_to = { $regex: sentTo, $options: "i" };
      }
      if (sentDate) {
        query.letter_sent_date = sentDate;
      }
      if (status) {
        query.status = status;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalLetters = await Letter.countDocuments(query);

      const totalPages = Math.ceil(totalLetters / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findLetter = await Letter.find(query)
        .sort({ createdAt: sortBy })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname",
        });

      if (!findLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no forwarded letters",
          Message_am: "ምንም የተላለፉ/የተመሩ ደብዳቤዎች የሎዎትም።",
        });
      }

      const lstOfForwardLtr = [];

      for (const items of findLetter) {
        const findFrwdLtr = await ForwardLetter.findOne({
          letter_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findFrwdLtr) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseForwarded: caseFind };

        lstOfForwardLtr.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        letters: lstOfForwardLtr,
        totalLetters,
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

const getForwardedLetterCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDLETTERSCC_API;
    const actualAPIKey = req?.headers?.get_frwdletterscc_api;
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

      const forwardedLetters = await ForwardLetter.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!forwardedLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd letters",
          Message_am: "ምንም የተላለፉ/የተመሩ ደብዳቤዎች የሎዎትም።",
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
      let letterNum = req?.query?.letter_number || "";
      let nimera = req?.query?.nimera || "";
      let letterType = req?.query?.letter_type || "";
      let sentFrom = req?.query?.sent_from || "";
      let sentTo = req?.query?.sent_to || "";
      let sentDate = req?.query?.letter_sent_date || "";
      let status = req?.query?.status || "";

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
      if (letterType === "" || letterType === null) {
        letterType = "";
      }

      const letterIds = forwardedList?.map((fwdLetter) => fwdLetter?.letter_id);

      const query = {};

      if (letterNum) {
        query.letter_number = { $regex: letterNum, $options: "i" };
      }
      if (nimera) {
        query.nimera = { $regex: nimera, $options: "i" };
      }
      if (letterType) {
        query.letter_type = letterType;
      }
      if (sentFrom) {
        query.sent_from = { $regex: sentFrom, $options: "i" };
      }
      if (sentTo) {
        query.sent_to = { $regex: sentTo, $options: "i" };
      }
      if (sentDate) {
        query.letter_sent_date = sentDate;
      }
      if (status) {
        query.status = status;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalLetters = await Letter.countDocuments(query);

      const totalPages = Math.ceil(totalLetters / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findLetter = await Letter.find(query)
        .sort({ createdAt: sortBy })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname",
        });

      if (!findLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd letters",
          Message_am: "ምንም የተላለፉ/የተመሩ (CC) ደብዳቤዎች የሎዎትም።",
        });
      }

      return res.status(StatusCodes.OK).json({
        letters: findLetter,
        totalLetters,
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

const getLetterForwardPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_LETTERFRWDPATH_API;
    const actualAPIKey = req?.headers?.get_letterfrwdpath_api;
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

      const letter_id = req?.params?.letter_id;

      if (!letter_id || !mongoose.isValidObjectId(letter_id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findLetter = await Letter.findOne({ _id: letter_id });
      const findForwardLetter = await ForwardLetter.findOne({
        letter_id: letter_id,
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

      if (!findLetter || !findForwardLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its forwards were not found",
          Message_am: "ደብዳቤው እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const forwardLetters = findForwardLetter?.path;

      return res.status(StatusCodes.OK).json({
        forwardDocs: forwardLetters,
        forwardId: findForwardLetter?._id,
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

// CONTINUE
const printForwardLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_LETTERPATHPRT_API;
    const actualAPIKey = req?.headers?.get_letterpathprt_api;
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

      const findForwardLetter = await ForwardLetter.findOne({ _id: id });

      const findLetter = await Letter.findOne({
        _id: findForwardLetter?.letter_id,
      });

      if (!findLetter || !findForwardLetter) {
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

      const forwardToPrint = findForwardLetter?.path?.find(
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
          Message_en: `The person who forwarded this letter (Letter Number: ${findLetter?.letter_number}) is not found among the office administrators or the archival.`,
          Message_am: `ይህንን ደብዳቤ ያስተላለፈው ሰው (የደብዳቤ ቁጥር፡ ${findLetter?.letter_number}) ከቢሮ ወይም መዝገብ ከቤት አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser?.findOne({
        _id: forwardToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received this letter (Letter Number: ${findLetter?.letter_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ደብዳቤ የተቀበለው ሰው (የደብዳቤ ቁጥር፡ ${findLetter?.letter_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      let checkWho = "offUsr";

      if (findArchivalSender) {
        checkWho = "arcUsr";
      }

      const sent_date = caseSubDate(forwardToPrint?.forwarded_date);
      const case_num = findLetter?.letter_number;
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
        "ForwardLetterPrint",
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
        await appendForwardLetterPrint(inputPath, text, outputPath);
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
  archivalLetterForward,
  officerLetterForward,
  getForwardedLetter,
  getForwardedLetterCC,
  getLetterForwardPath,
  printForwardLetter,
};
