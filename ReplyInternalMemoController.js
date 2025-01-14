const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const Notification = require("../../model/Notifications/Notification");
const InternalMemo = require("../../model/InternalMemo/InternalMemo");
const ReplyInternalMemo = require("../../model/ReplyInternalMemo/ReplyInternalMemo");
const ForwardInternalMemo = require("../../model/ForwardInternalMemo/ForwardInternalMemo");
const ReplyInternalMemoHistory = require("../../model/ReplyInternalMemo/ReplyInternalMemoHistory");

const { join } = require("path");
const mongoose = require("mongoose");
const formidable = require("formidable");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile } = require("fs/promises");
const {
  appendReplyInternalMemoPrint,
} = require("../../middleware/replyInternalMemoPrt");

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

const replyInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CRRPLYINTMEM_API;
    const actualAPIKey = req?.headers?.get_crrplyintmem_api;
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
      const internal_memo_id = req?.params?.internal_memo_id;
      const onlineUserList = global?.onlineUserList;

      if (!internal_memo_id || !mongoose.isValidObjectId(internal_memo_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the internal memo you want to reply to",
          Message_am: "እባክዎን መልስ መስጠት የሚፈልጉትን የዉስጥ ማስታወሻ ደብዳቤ ያቅርቡ",
        });
      }

      const findInternalMemo = await InternalMemo.findOne({
        _id: internal_memo_id,
      });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en:
            "The internal memo that you want to place your reply is not found",
          Message_am: "መልስዎን ማስቀመጥ የሚፈልጉበት የዉስጥ ማስታወሻ ደብዳቤ አልተገኘም",
        });
      }

      const findForwardInternalMemo = await ForwardInternalMemo.findOne({
        internal_memo_id: internal_memo_id,
      });

      if (!findForwardInternalMemo) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Since the internal memo has not been forwarded yet, replies cannot be sent for this memo",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤው እስካሁን ስላልተላለፈ ለዚህ ደብዳቤ ምላሽ መላክ አይቻልም",
        });
      }

      const findReplyingPersonInForward = findForwardInternalMemo?.path?.filter(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (
        findReplyingPersonInForward?.length === 0 &&
        findInternalMemo?.createdBy?.toString() !== requesterId?.toString()
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You cannot reply to this internal memo as it has not been forwarded to you or created by you",
          Message_am:
            "ይህ የዉስጥ ማስታወሻ ደብዳቤ ለእርስዎ አልተላከም ወይም በእርስዎ አልተፈጠርም (not created) ፤ ስለዚህ ለዚህ ደብዳቤ መልስ መስጠት አይችሉም",
        });
      }

      if (
        findReplyingPersonInForward?.length > 0 &&
        findInternalMemo?.createdBy?.toString() !== requesterId?.toString()
      ) {
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

      const findReplyInternalMemo = await ReplyInternalMemo.findOne({
        internal_memo_id: internal_memo_id,
      });

      if (findReplyInternalMemo) {
        const findNormalForward = findReplyingPersonInForward?.find(
          (item) => item?.cc === "no"
        );

        const findSendingUserInReply = findReplyInternalMemo?.path?.filter(
          (item) => item?.to?.toString() === requesterId?.toString()
        );

        if (findSendingUserInReply?.length > 0) {
          const findNormal = findSendingUserInReply?.find(
            (item) => item?.cc === "no"
          );

          if (
            !findNormal &&
            !findNormalForward &&
            findInternalMemo?.createdBy?.toString() !== requesterId?.toString()
          ) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en:
                "You cannot reply to this internal memo as it was only CC'd to you",
              Message_am:
                "ለእርስዎ CC ብቻ ስለነበር የተደረገው ለዚህ የዉስጥ ማስታወሻ ደብዳቤ መልስ መስጠት አይችሉም",
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

        const findCCForwardTo = findForwardInternalMemo?.path?.filter(
          (item) => item?.to?.toString() === to?.toString()
        );

        if (
          cc === "no" &&
          findCCForwardTo?.length === 0 &&
          findInternalMemo?.createdBy?.toString() !== to?.toString()
        ) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: `You cannot reply to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as this internal memo was not forwarded to him/her`,
            Message_am: `ይህ የዉስጥ ማስታወሻ ወደ እሱ/ሷ ስላልተላከ ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ምላሽ መስጠት አይችሉም`,
          });
        }

        if (
          cc === "no" &&
          findCCForwardTo?.length > 0 &&
          findInternalMemo?.createdBy?.toString() !== to?.toString()
        ) {
          const findNormal = findCCForwardTo?.find((item) => item?.cc === "no");

          if (!findNormal) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en: "You cannot reply to a person who is only CC'd",
              Message_am: "CC ብቻ ለተደረገለት ሰው መልስ መስጠት አይችሉም",
            });
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
            "InternalMemoReplyFiles",
            uniqueSuffix + "-" + attachment?.originalFilename
          );

          attachmentName = uniqueSuffix + "-" + attachment?.originalFilename;

          await writeFile(path, letterReplyAttachmentBuffer);
        }

        if (!findReplyInternalMemo) {
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

          const createReplyInternalMemo = await ReplyInternalMemo.create({
            internal_memo_id: internal_memo_id,
            forward_internal_memo_id: findForwardInternalMemo?._id,
            path,
          });

          const updateHistory = [
            {
              updatedByOfficeUser: requesterId,
              action: "create",
            },
          ];

          try {
            await ReplyInternalMemoHistory.create({
              reply_internal_memo_id: createReplyInternalMemo?._id,
              updateHistory,
              history: createReplyInternalMemo?.toObject(),
            });
          } catch (error) {
            console.log(
              `Reply history for internal memo with ID (${findInternalMemo?._id}) is not created`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for an internal memo is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `An internal memo is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `ለዉስጥ ማስታወሻ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
                : `ለዉስጥ ማስታወሻ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
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
            io.to(user?.socketID).emit("internal_memo_reply_notification", {
              Message_en: `Replies for internal memo is sent to you.`,
              Message_am: `የዉስጥ ማስታወሻ ደብዳቤ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.CREATED).json({
            Message_en: `Reply for internal memo is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለዉስጥ ማስታወሻ ደብዳቤ ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
          });
        }

        if (findReplyInternalMemo) {
          const updateInternalMemoReply =
            await ReplyInternalMemo.findOneAndUpdate(
              {
                _id: findReplyInternalMemo?._id,
              },
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

          if (!updateInternalMemoReply) {
            return res.status(StatusCodes.EXPECTATION_FAILED).json({
              Message_en: `The attempt to reply to an memo letter was not successful.`,
              Message_am: `ለዉስጥ ማስታወሻ ደብዳቤዉ መልስ ለመስጠት የተደረገው ሙከራ አልተሳካም።`,
            });
          }

          try {
            await ReplyInternalMemoHistory.findOneAndUpdate(
              {
                reply_internal_memo_id: findReplyInternalMemo?._id,
              },
              {
                $push: {
                  updateHistory: {
                    updatedByOfficeUser: requesterId,
                    action: "update",
                  },
                  history: updateInternalMemoReply?.toObject(),
                },
              }
            );
          } catch (error) {
            console.log(
              `Reply history for internal memo with ID (${findInternalMemo?._id}) is not updated successfully`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for an internal memo is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `An internal memo is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `ለዉስጥ ማስታወሻ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
                : `ለዉስጥ ማስታወሻ ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል።`,
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
            io.to(user?.socketID).emit("internal_memo_reply_notification", {
              Message_en: `Replies for internal memo is sent to you.`,
              Message_am: `የዉስጥ ማስታወሻ ደብዳቤ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.OK).json({
            Message_en: `Reply for internal memo is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለዉስጥ ማስታወሻ ደብዳቤ ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
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

const getRepliedInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYINTMEMS_API;
    const actualAPIKey = req?.headers?.get_rplyintmems_api;
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

      const repliedInternalMemos = await ReplyInternalMemo.find({
        "path.to": requesterId,
        "path.cc": "no",
      });

      if (!repliedInternalMemos) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied internal memos",
          Message_am: "ምንም መልሶች የተሰጡባቸው የዉስጥ ማስታወሻ ደብዳቤዎች የሎዎትም",
        });
      }

      for (const replyPath of repliedInternalMemos) {
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

      const letterIds = replyList?.map((rplyLtr) => rplyLtr?.internal_memo_id);

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

      const lstOfReplyInternalMemo = [];

      for (const items of findInternalMemo) {
        const findReplyInternalMem = await ReplyInternalMemo.findOne({
          internal_memo_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findReplyInternalMem) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseReplied: caseFind };

        lstOfReplyInternalMemo.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        internalMemos: lstOfReplyInternalMemo,
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

const getRepliedInternalMemoCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYINTMEMSCC_API;
    const actualAPIKey = req?.headers?.get_rplyintmemscc_api;
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

      const repliedInternalMemos = await ReplyInternalMemo.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!repliedInternalMemos) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no CC'd internal memos",
          Message_am: "ምንም CC የተደረጉ የዉስጥ ማስታወሻ ደብዳቤዎች የሎዎትም",
        });
      }

      for (const replyPath of repliedInternalMemos) {
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

      const letterIds = replyList?.map((rplyLtr) => rplyLtr?.internal_memo_id);

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

const getInternalMemoRepliedPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYINTMEMSPATH_API;
    const actualAPIKey = req?.headers?.get_rplyintmemspath_api;
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

      const findInternalMemo = await InternalMemo.findOne({
        _id: internal_memo_id,
      });
      const findReplyInternalMemo = await ReplyInternalMemo.findOne({
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

      if (!findInternalMemo || !findReplyInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its replies were not found",
          Message_am: "ደብዳቤው እና የምላሾቹ ዝርዝሮች አልተገኙም",
        });
      }

      const replyLetters = findReplyInternalMemo?.path;

      return res.status(StatusCodes.OK).json({
        repliedDocs: replyLetters,
        replyId: findReplyInternalMemo?._id,
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

const printReplyInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRTRPLYINTMEMS_API;
    const actualAPIKey = req?.headers?.get_prtrplyintmems_api;
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

      const findReplyInternalMemo = await ReplyInternalMemo.findOne({
        _id: id,
      });
      const findInternalMemo = await InternalMemo.findOne({
        _id: findReplyInternalMemo?.internal_memo_id,
      });

      if (!findInternalMemo || !findReplyInternalMemo) {
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

      const replyToPrint = findReplyInternalMemo?.path?.find(
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
          Message_en: `The person who replied to this internal memo is not found among the office administrators.`,
          Message_am: `ለዚህ የዉስጥ ማስታወሻ ደብዳቤ ምላሽ የሰጡት ተጠቃሚ በቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser.findOne({
        _id: replyToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received the answer to this internal memo is not found among the office administrators.`,
          Message_am: `ይህንን የዉስጥ ማስታወሻ ደብዳቤ መልስ የተቀበለው ተጠቃሚ ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const sent_date = caseSubDate(replyToPrint?.replied_date);
      const case_num = "ሲስተም-ID: " + findInternalMemo?._id;
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
        "ReplyInternalMemoPrint",
        uniqueSuffix + "-replyinternalletter.pdf"
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
        await appendReplyInternalMemoPrint(inputPath, text, outputPath);
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
  replyInternalMemo,
  getRepliedInternalMemo,
  getRepliedInternalMemoCC,
  getInternalMemoRepliedPath,
  printReplyInternalMemo,
};
