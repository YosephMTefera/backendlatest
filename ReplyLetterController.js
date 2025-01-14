const Letter = require("../../model/Letters/Letter");
const Division = require("../../model/Divisions/Divisions");
const ReplyLetter = require("../../model/ReplyLetters/ReplyLetters");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const ForwardLetter = require("../../model/ForwardLetters/ForwardLetter");
const Notification = require("../../model/Notifications/Notification");
const ReplyLetterHistory = require("../../model/ReplyLetters/ReplyLettersHistory");

const { join } = require("path");
const mongoose = require("mongoose");
const formidable = require("formidable");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile } = require("fs/promises");
const { appendReplyLetterPrint } = require("../../middleware/replyLetterPrint");

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

const replyLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYLTRS_API;
    const actualAPIKey = req?.headers?.get_rplyltrs_api;
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
      const letter_id = req?.params?.letter_id;
      const onlineUserList = global?.onlineUserList;

      if (!letter_id || !mongoose.isValidObjectId(letter_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the letter you want to reply to",
          Message_am: "እባክዎን መልስ መስጠት የሚፈልጉትን ደብዳቤ ያቅርቡ",
        });
      }

      const findLetter = await Letter.findOne({ _id: letter_id });

      if (!findLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en:
            "The letter that you want to place your reply is not found",
          Message_am: "መልስዎን ማስቀመጥ የሚፈልጉበት ደብዳቤ አልተገኘም",
        });
      }

      if (findLetter?.status === "created") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter has not been forwarded, so replies cannot be given.",
          Message_am: "ደብዳቤው አልተላከም ስለዚህ ምንም አይነት ምላሽ/reply ማቅረብ አይችሉም።",
        });
      }

      const findForwardLetter = await ForwardLetter.findOne({
        letter_id: letter_id,
      });

      if (!findForwardLetter) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Since the letter has not been forwarded yet, replies cannot be sent for this letter",
          Message_am: "ደብዳቤው እስካሁን ስላልተላለፈ ለዚህ ደብዳቤ ምላሽ መላክ አይቻልም",
        });
      }

      const findReplyingPersonInForward = findForwardLetter?.path?.filter(
        (item) =>
          item?.to?.toString() === requesterId?.toString() ||
          item?.from_office_user?.toString() === requesterId?.toString()
      );

      if (findReplyingPersonInForward?.length === 0) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You cannot reply to this letter as it has not been forwarded to you or forwarded by you",
          Message_am:
            "ይህ ደብዳቤ ለእርስዎ አልተላከም ወይም በእርስዎ አልተላከም ፤ ስለዚህ ለዚህ ደብዳቤ መልስ መስጠት አይችሉም",
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

      const findReplyLetter = await ReplyLetter.findOne({
        letter_id: letter_id,
      });

      if (findReplyLetter) {
        const findNormalForward = findReplyingPersonInForward?.find(
          (item) => item?.cc === "no"
        );
        const findSendingUserInReply = findReplyLetter?.path?.filter(
          (item) =>
            item?.to?.toString() === requesterId?.toString() ||
            item?.from_office_user?.toString() === requesterId?.toString()
        );

        if (findSendingUserInReply?.length > 0) {
          const findNormal = findSendingUserInReply?.find(
            (item) => item?.cc === "no"
          );

          if (!findNormal && !findNormalForward) {
            return res.status(StatusCodes.FORBIDDEN).json({
              Message_en:
                "You cannot reply to this letter as it was only CC'd to you",
              Message_am: "ለእርስዎ CC ብቻ ስለነበር የተደረገው ለዚህ ደብዳቤ መልስ መስጠት አይችሉም",
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

        const findCCForwardTo = findForwardLetter?.path?.filter(
          (item) =>
            item?.to?.toString() === to?.toString() ||
            item?.from_office_user?.toString() === to?.toString()
        );

        if (cc === "no" && findCCForwardTo?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: `You cannot reply to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as this letter was not forwarded to him/her`,
            Message_am: `ይህ ደብዳቤ ወደ እሱ/ሷ ስላልተላከ ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ምላሽ መስጠት አይችሉም`,
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
                "Invalid attachment letter format please try again. Only accepts '.pdf'.",
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
            "LetterReplyFiles",
            uniqueSuffix + "-" + attachment?.originalFilename
          );

          attachmentName = uniqueSuffix + "-" + attachment?.originalFilename;

          await writeFile(path, letterReplyAttachmentBuffer);
        }

        if (!findReplyLetter) {
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

          const createReply = await ReplyLetter.create({
            letter_id: letter_id,
            forward_letter_id: findForwardLetter?._id,
            path,
          });

          const updateHistory = [
            {
              updatedByOfficeUser: requesterId,
              action: "create",
            },
          ];

          try {
            await ReplyLetterHistory.create({
              reply_letter_id: createReply?._id,
              updateHistory,
              history: createReply?.toObject(),
            });
          } catch (error) {
            console.log(
              `Reply history for letter with letter number (${findLetter?.letter_number}) is not created`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for letter with letter number ${findLetter?.letter_number} is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Letter with letter number ${findLetter?.letter_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `የደብዳቤ ቁጥር ${findLetter?.letter_number} ባለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
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
            io.to(user?.socketID).emit("letter_reply_notification", {
              Message_en: `Replies for letter with letter number ${findLetter?.letter_number} is sent to you.`,
              Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ላለው ደብዳቤ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.CREATED).json({
            Message_en: `Reply for letter with letter number ${findLetter?.letter_number} is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለደብዳቤ ቁጥር ${findLetter?.letter_number} ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
          });
        }

        if (findReplyLetter) {
          const updateReply = await ReplyLetter.findOneAndUpdate(
            { _id: findReplyLetter?._id },
            {
              $push: {
                path: {
                  replied_date: new Date(),
                  from_office_user: requesterId,
                  to: to,
                  cc: cc,
                  attachment: attachmentName,
                  remark: remark,
                },
              },
            },
            { new: true }
          );

          if (!updateReply) {
            return res.status(StatusCodes.EXPECTATION_FAILED).json({
              Message_en: `The attempt to reply to letter number ${findLetter?.letter_number} for ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} was unsuccessful.`,
              Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}ን ለመመለስ የተደረገው ሙከራ አልተሳካም።`,
            });
          }

          try {
            await ReplyLetterHistory.findOneAndUpdate(
              { reply_letter_id: findReplyLetter?._id },
              {
                $push: {
                  updateHistory: {
                    updatedByOfficeUser: requesterId,
                    action: "update",
                  },
                  history: updateReply?.toObject(),
                },
              }
            );
          } catch (error) {
            console.log(
              `Reply history of letter with letter number ${findLetter?.letter_number} is not updated successfully`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for letter with letter number ${findLetter?.letter_number} is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Letter with letter number ${findLetter?.letter_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `የደብዳቤ ቁጥር ${findLetter?.letter_number} ባለው ደብዳቤ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
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
            io.to(user?.socketID).emit("letter_reply_notification", {
              Message_en: `Replies for letter with letter number ${findLetter?.letter_number} is sent to you.`,
              Message_am: `የደብዳቤ ቁጥር ${findLetter?.letter_number} ላለው ደብዳቤ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.OK).json({
            Message_en: `Reply for letter with letter number ${findLetter?.letter_number} is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለደብዳቤ ቁጥር ${findLetter?.letter_number} ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
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

const getRepliedLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYLTRSPATH_API;
    const actualAPIKey = req?.headers?.get_rplyltrspath_api;
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

      const repliedLetters = await ReplyLetter.find({
        "path.to": requesterId,
        "path.cc": "no",
      });

      if (!repliedLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied letters",
          Message_am: "ምንም መልሶች የተሰጡባቸው ደብዳቤዎች የሎዎትም",
        });
      }

      for (const replyPath of repliedLetters) {
        const pathToRequester = replyPath?.path?.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "no"
        );

        if (pathToRequester) {
          replyList?.push(replyPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || 1;
      let status = req?.query?.status || "";
      let letterNum = req?.query?.letter_number || "";
      let nimera = req?.query?.nimera || "";
      let letterType = req?.query?.letter_type || "";
      let sentFrom = req?.query?.sent_from || "";
      let sentTo = req?.query?.sent_to || "";
      let sentDate = req?.query?.letter_sent_date || "";

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

      const letterIds = replyList?.map((fwdLetter) => fwdLetter?.letter_id);

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

      const findLetters = await Letter.find(query)
        .sort({ createdAt: sortBy })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname",
        });

      if (!findLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied letters",
          Message_am: "ምንም መልሶች የተሰጡባቸው ደብዳቤዎች የሎዎትም",
        });
      }

      const lstOfReplyLtr = [];

      for (const items of findLetters) {
        const findReplyLtr = await ReplyLetter.findOne({
          letter_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findReplyLtr) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseReplied: caseFind };

        lstOfReplyLtr.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        letters: lstOfReplyLtr,
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

const getRepliedLetterCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYLTRSPATHCC_API;
    const actualAPIKey = req?.headers?.get_rplyltrspathcc_api;
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

      const repliedLetters = await ReplyLetter.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!repliedLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied (CC) letters",
          Message_am: "ምንም መልሶች የተሰጡባቸው (CC) ደብዳቤዎች የሎዎትም",
        });
      }

      for (const replyPath of repliedLetters) {
        const pathToRequester = replyPath?.path?.find(
          (path) =>
            path?.to?.toString() === requesterId?.toString() &&
            path?.cc === "yes"
        );

        if (pathToRequester) {
          replyList?.push(replyPath);
        }
      }

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || 1;
      let status = req?.query?.status || "";
      let letterNum = req?.query?.letter_number || "";
      let nimera = req?.query?.nimera || "";
      let letterType = req?.query?.letter_type || "";
      let sentFrom = req?.query?.sent_from || "";
      let sentTo = req?.query?.sent_to || "";
      let sentDate = req?.query?.letter_sent_date || "";

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

      const letterIds = replyList?.map((fwdLetter) => fwdLetter?.letter_id);

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

      const findLetters = await Letter.find(query)
        .sort({ createdAt: sortBy })
        .skip(skip)
        .limit(limit)
        .populate({
          path: "createdBy",
          select: "_id firstname middlename lastname",
        });

      if (!findLetters) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied (CC) letters",
          Message_am: "ምንም መልሶች የተሰጡባቸው (CC) ደብዳቤዎች የሎዎትም",
        });
      }

      return res.status(StatusCodes.OK).json({
        letters: findLetters,
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

const getReplyLetterPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYPATHLTR_API;
    const actualAPIKey = req?.headers?.get_rplypathltr_api;
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

      const letter_id = req?.params?.letter_id;

      if (!letter_id || !mongoose.isValidObjectId(letter_id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findLetter = await Letter.findOne({ _id: letter_id });
      const findReplyLetter = await ReplyLetter.findOne({
        letter_id: letter_id,
      })
        .populate({
          path: "path.from_office_user",
          select: "_id firstname middlename lastname position username level",
        })
        .populate({
          path: "path.to",
          select: "_id firstname middlename lastname position username level",
        });

      if (!findLetter || !findReplyLetter) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The letter and its replies were not found",
          Message_am: "ደብዳቤው እና ምላሾቹ አልተገኙም",
        });
      }

      const replyLetters = findReplyLetter?.path;

      return res.status(StatusCodes.OK).json({
        repliedDocs: replyLetters,
        replyId: findReplyLetter?._id,
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

const printReplyLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYPRTLTR_API;
    const actualAPIKey = req?.headers?.get_rplyprtltr_api;
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

      const findReplyLetter = await ReplyLetter.findOne({ _id: id });

      const findLetter = await Letter.findOne({
        _id: findReplyLetter?.letter_id,
      });

      if (!findLetter || !findReplyLetter) {
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

      const replyToPrint = findReplyLetter?.path?.find(
        (reply) => reply?._id?.toString() === reply_forward_id?.toString()
      );

      if (!replyToPrint) {
        return res.status(StatusCodes.CONFLICT).json({
          Message_en: "The specific reply path does not exist",
          Message_am: "የተወሰነው የምላሽ መንገድ የለም",
        });
      }

      const findSenderUser = await OfficeUser.findOne({
        _id: replyToPrint?.from_office_user,
      });

      if (!findSenderUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who replied to this letter (Letter Number: ${findLetter?.letter_number}) is not found among the office administrators.`,
          Message_am: `ለዚህ ደብዳቤ ምላሽ የሰጡት ሰው (የደብዳቤ ቁጥር፡ ${findLetter?.letter_number}) በቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser.findOne({
        _id: replyToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received the answer to this letter (Letter Number: ${findLetter?.letter_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ደብዳቤ መልስ የተቀበለው ሰው (የደብዳቤ ቁጥር፡ ${findLetter?.letter_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const sent_date = caseSubDate(replyToPrint?.replied_date);
      const case_num = findLetter?.letter_number;
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
        "ReplyLetterPrint",
        uniqueSuffix + "-replyletter.pdf"
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
        await appendReplyLetterPrint(inputPath, text, outputPath);

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
  replyLetter,
  getRepliedLetter,
  getRepliedLetterCC,
  getReplyLetterPath,
  printReplyLetter,
};
