const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const Notification = require("../../model/Notifications/Notification");
const InternalMemo = require("../../model/InternalMemo/InternalMemo");
const ForwardInternalMemo = require("../../model/ForwardInternalMemo/ForwardInternalMemo");
const ForwardInternalMemoHistory = require("../../model/ForwardInternalMemo/ForwardInternalMemoHistory");

const { join } = require("path");
const mongoose = require("mongoose");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const {
  appendForwardInternalMemoPrint,
} = require("../../middleware/forwardInternalMemoPrt");

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

const officerInternalMemoForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CRFRWDINTMEM_API;
    const actualAPIKey = req?.headers?.get_crfrwdintmem_api;
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
      const internal_memo_id = req?.body?.internal_memo_id;
      const forwardArray = req?.body?.forwardArray;
      const onlineUserList = global?.onlineUserList;

      if (!internal_memo_id || !mongoose.isValidObjectId(internal_memo_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please provide the internal memo to be forwarded",
          Message_am: "እባክዎ የሚላከውን የዉስጥ ማስታወሻ ደብዳቤ ያቅርቡ",
        });
      }

      const findInternalMemo = await InternalMemo.findOne({
        _id: internal_memo_id,
      });
      const findForwardInternalMemo = await ForwardInternalMemo.findOne({
        internal_memo_id: internal_memo_id,
      });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The internal memo to be forwarded/sent is not found`,
          Message_am: `የሚተላለፈው/የሚላከው የዉስጥ ማስታወሻ ደብዳቤ አልተገኘም`,
        });
      }

      const isUserSender = findForwardInternalMemo?.path?.some(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (findForwardInternalMemo) {
        if (
          !isUserSender &&
          findInternalMemo?.createdBy?.toString() !== requesterId?.toString()
        ) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You are unauthorized to forward this letter as it was not initially forwarded to you",
            Message_am: "መጀመሪያ ወደ እርስዎ ስላልተላከ ይህንን ደብዳቤ ማስተላለፍ አልተፈቀደልዎትም",
          });
        }

        const checkIfUserIsCC = findForwardInternalMemo?.path?.filter(
          (item) => item?.to?.toString() === requesterId?.toString()
        );

        if (
          findInternalMemo?.createdBy?.toString() !== requesterId?.toString()
        ) {
          if (checkIfUserIsCC?.length > 0) {
            const findNormal = checkIfUserIsCC?.find(
              (item) => item?.cc === "no"
            );

            if (!findNormal) {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en:
                  "You cannot forward this internal memo as it was only CC'd to you, not directly forwarded",
                Message_am:
                  "የዉስጥ ማስታወሻ ደብዳቤዉ ፤ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ ማስተላለፍ/መላክ አይችሉም",
              });
            }
          }
        }
      }

      if (
        !findForwardInternalMemo &&
        findInternalMemo?.createdBy?.toString() !== requesterId?.toString()
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

          const isUserRecipient = findForwardInternalMemo?.path?.filter(
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

          const findForwardInternalMemoCheck =
            await ForwardInternalMemo.findOne({
              internal_memo_id: internal_memo_id,
            });

          if (findForwardInternalMemoCheck) {
            const forwardInternalMemoToOfficer =
              await ForwardInternalMemo.findOneAndUpdate(
                { internal_memo_id: internal_memo_id },
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

            if (!forwardInternalMemoToOfficer) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The internal memo is not successfully forwarded to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
                Message_am: `የዉስጥ ማስታወሻ ደብዳቤው ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} አልተላለፈም/አልተላከም።`,
              });
            }

            try {
              await ForwardInternalMemoHistory.findOneAndUpdate(
                {
                  forward_internal_memo_id: forwardInternalMemoToOfficer?._id,
                },
                {
                  $push: {
                    updateHistory: {
                      updatedByOfficeUser: requesterId,
                      action: "update",
                    },
                    history: forwardInternalMemoToOfficer?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Forward history for internal memo with ID ${findInternalMemo?._id} is not updated successfully`
              );
            }
          }

          if (!findForwardInternalMemoCheck) {
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

            const newInternalMemoForward = await ForwardInternalMemo.create({
              internal_memo_id: internal_memo_id,
              path: path,
            });

            const updateHistory = [
              {
                updatedByOfficeUser: requesterId,
                action: "create",
              },
            ];

            try {
              await ForwardInternalMemoHistory.create({
                forward_internal_memo_id: newInternalMemoForward?._id,
                updateHistory: updateHistory,
                history: newInternalMemoForward?.toObject(),
              });
            } catch (error) {
              console.log(
                `Forward history for internal memo with ID: ${findInternalMemo?._id} is not created successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en:
              singlePathSend?.cc === "no"
                ? `Internal memo is forwarded to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Internal memo is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              singlePathSend?.cc === "no"
                ? `የወስጥ ማስታወሻ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ ተላልፏል/ተልኳል።`
                : `የወስጥ ማስታወሻ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: findAcceptorUser?._id,
            notifcation_type: "InternalMemo",
            document_id: findInternalMemo?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(findAcceptorUser?._id, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("internal_memo_forward_notification", {
              Message_en: `Internal memo is forwarded to you.`,
              Message_am: `የዉስጥ ማስታወሻ ደብዳቤው ወደ እርስዎ ተልኳል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en:
            "Internal memo is forwarded successfully to the recipients.",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤው ለተቀባዮች በተሳካ ሁኔታ ተላልፏል/ተልኳል።",
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

const getInternalMemoForward = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDGETINTMEMOS_API;
    const actualAPIKey = req?.headers?.get_frwdgetintmemos_api;
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

      const forwardedLetters = await ForwardInternalMemo.find({
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
      let status = req?.query?.status || "";
      let subject = req?.query?.subject || "";
      let createdBy = req?.query?.createdBy || "";
      let verifiedBy = req?.query?.verified_by || "";
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
      if (!createdBy) {
        createdBy = null;
      }
      if (!verifiedBy) {
        verifiedBy = null;
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

      const findOfficerCreatedBy = await OfficeUser.findOne({ _id: createdBy });
      if (!findOfficerCreatedBy) {
        createdBy = "";
      }

      const findOfficerVerifiedBy = await OfficeUser.findOne({
        _id: verifiedBy,
      });
      if (!findOfficerVerifiedBy) {
        verifiedBy = "";
      }

      const letterIds = forwardedList?.map(
        (fwdLtr) => fwdLtr?.internal_memo_id
      );

      const query = {};

      if (subject) {
        query.subject = { $regex: subject, $options: "i" };
      }
      if (status) {
        query.status = status;
      }
      if (createdBy) {
        query.createdBy = createdBy;
      }
      if (verifiedBy) {
        query.verified_by = verifiedBy;
      }
      if (verifiedDate) {
        query.verified_date = verifiedDate;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalInternalMemos = await InternalMemo.countDocuments(query);

      const totalPages = Math.ceil(totalInternalMemos / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findInternalMemo = await InternalMemo.find(query)
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
          path: "verified_by",
          select: "_id firstname middlename lastname username position",
        })
        .populate({
          path: "updated_by.update_officer",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "to_whom.internal_office",
          select: "_id firstname middlename lastname username position",
        })
        .populate({
          path: "internal_cc.internal_office",
          select: "_id firstname middlename lastname position",
        });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal memos are not found",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤዎች አልተገኙም",
        });
      }

      const lstOfForwardedInternalMemo = [];

      for (const items of findInternalMemo) {
        const findFrwdInternalMemo = await ForwardInternalMemo.findOne({
          internal_memo_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findFrwdInternalMemo) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseForwarded: caseFind };

        lstOfForwardedInternalMemo.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        internalMemos: lstOfForwardedInternalMemo,
        totalInternalMemos,
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

const getInternalMemoForwardCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDGETINTMEMOSCC_API;
    const actualAPIKey = req?.headers?.get_frwdgetintmemoscc_api;
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

      const forwardedLetters = await ForwardInternalMemo.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!forwardedLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd letters",
          Message_am: "ምንም ግልባጭ የተደረጉ ደብዳቤዎች የሎዎትም።",
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
      let status = req?.query?.status || "";
      let subject = req?.query?.subject || "";
      let createdBy = req?.query?.createdBy || "";
      let verifiedBy = req?.query?.verified_by || "";
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
      if (!createdBy) {
        createdBy = null;
      }
      if (!verifiedBy) {
        verifiedBy = null;
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

      const findOfficerCreatedBy = await OfficeUser.findOne({ _id: createdBy });
      if (!findOfficerCreatedBy) {
        createdBy = "";
      }

      const findOfficerVerifiedBy = await OfficeUser.findOne({
        _id: verifiedBy,
      });
      if (!findOfficerVerifiedBy) {
        verifiedBy = "";
      }

      const letterIds = forwardedList?.map(
        (fwdLtr) => fwdLtr?.internal_memo_id
      );

      const query = {};

      if (subject) {
        query.subject = { $regex: subject, $options: "i" };
      }
      if (status) {
        query.status = status;
      }
      if (createdBy) {
        query.createdBy = createdBy;
      }
      if (verifiedBy) {
        query.verified_by = verifiedBy;
      }
      if (verifiedDate) {
        query.verified_date = verifiedDate;
      }
      if (letterIds) {
        query._id = { $in: letterIds };
      }

      const totalInternalMemos = await InternalMemo.countDocuments(query);

      const totalPages = Math.ceil(totalInternalMemos / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findInternalMemo = await InternalMemo.find(query)
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
          path: "verified_by",
          select: "_id firstname middlename lastname username position",
        })
        .populate({
          path: "updated_by.update_officer",
          select: "_id firstname middlename lastname position",
        })
        .populate({
          path: "to_whom.internal_office",
          select: "_id firstname middlename lastname username position",
        })
        .populate({
          path: "internal_cc.internal_office",
          select: "_id firstname middlename lastname position",
        });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal memos are not found",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤዎች አልተገኙም",
        });
      }

      return res.status(StatusCodes.OK).json({
        internalMemos: findInternalMemo,
        totalInternalMemos,
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

const getInternalMemoForwardedPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDGETINTMEMPATH_API;
    const actualAPIKey = req?.headers?.get_frwdgetintmempath_api;
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

      const internal_memo_id = req?.params?.internal_memo_id;

      if (!internal_memo_id || !mongoose.isValidObjectId(internal_memo_id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findInternalMemos = await InternalMemo.findOne({
        _id: internal_memo_id,
      });
      const findForwardInternalMemo = await ForwardInternalMemo.findOne({
        internal_memo_id: internal_memo_id,
      })
        .populate({
          path: "path.from_office_user",
          select: "_id firstname middlename lastname position username level",
        })
        .populate({
          path: "path.to",
          select: "_id firstname middlename lastname position username level",
        });

      if (!findInternalMemos || !findForwardInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its forwards were not found",
          Message_am: "ደብዳቤው እና የተመራባቸው/የሄደባቸው የተጠቃሚዋች ዝርዝር አልተገኙም",
        });
      }

      const forwardLetters = findForwardInternalMemo?.path;

      return res.status(StatusCodes.OK).json({
        forwardDocs: forwardLetters,
        forwardId: findForwardInternalMemo?._id,
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

const printForwardInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_FRWDPRTINTMEMPATH_API;
    const actualAPIKey = req?.headers?.get_frwdprtintmempath_api;
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

      const findForwardInternalMemo = await ForwardInternalMemo.findOne({
        _id: id,
      });
      const findInternalMemo = await InternalMemo.findOne({
        _id: findForwardInternalMemo?.internal_memo_id,
      });

      if (!findInternalMemo || !findForwardInternalMemo) {
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

      const forwardToPrint = findForwardInternalMemo?.path?.find(
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

      if (!findOfficeUserSender) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who forwarded this letter (internal memo letter) is not found among the office administrators.`,
          Message_am: `ይህንን ደብዳቤ ያስተላለፈው ሰው (የዉስጥ ማስታወሻ ደብዳቤ) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser?.findOne({
        _id: forwardToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received this letter (internal memo letter) is not found among the office administrators.`,
          Message_am: `ይህንን ደብዳቤ የተቀበለው ሰው (የዉስጥ ደብዳቤ) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const sent_date = caseSubDate(forwardToPrint?.forwarded_date);
      const case_num = "ሲስተም-ID: " + findInternalMemo?._id;
      const sent_from =
        findOfficeUserSender?.firstname +
        " " +
        findOfficeUserSender?.middlename +
        " " +
        findOfficeUserSender?.lastname;
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

      let signatureImg = "";
      if (findOfficeUserSender) {
        signatureImg = findOfficeUserSender?.signature;
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
        "ForwardInternalMemoPrint",
        uniqueSuffix + "-forwardinternalmemo.pdf"
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
        await appendForwardInternalMemoPrint(inputPath, text, outputPath);

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
  officerInternalMemoForward,
  getInternalMemoForward,
  getInternalMemoForwardCC,
  getInternalMemoForwardedPath,
  printForwardInternalMemo,
};
