const Case = require("../../model/Cases/Case");
const CaseList = require("../../model/CaseLists/CaseList");
const Division = require("../../model/Divisions/Divisions");
const ReplyCase = require("../../model/ReplyCases/ReplyCases");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const ForwardCase = require("../../model/ForwardCases/ForwardCase");
const Notification = require("../../model/Notifications/Notification");
const ReplyCaseHistory = require("../../model/ReplyCases/ReplyCasesHistory");

const { join } = require("path");
const mongoose = require("mongoose");
const formidable = require("formidable");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile } = require("fs/promises");
const { appendReplyCasePrint } = require("../../middleware/replyCasePrint");

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

const replyCase = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CAREPLYOF_API;
    const actualAPIKey = req?.headers?.get_careplyof_api;
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
      const case_id = req?.params?.case_id;
      const onlineUserList = global?.onlineUserList;

      if (!case_id || !mongoose.isValidObjectId(case_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the case you want to reply to",
          Message_am: "እባክዎን መልስ መስጠት የሚፈልጉትን ጉዳይ ያቅርቡ",
        });
      }

      const findCase = await Case.findOne({ _id: case_id });

      if (!findCase) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The case that you want to place your reply is not found",
          Message_am: "መልስዎን ማስቀመጥ የሚፈልጉበት ጉዳይ አልተገኘም",
        });
      }

      if (
        findCase?.status === "rejected" ||
        findCase?.status === "responded" ||
        findCase?.status === "verified"
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Cases that are rejected or responded/verified, cannot accept any replies",
          Message_am:
            "ውድቅ የተደረጉ ወይም ምላሽ የተሰጣቸው/የተረጋገጡ ጉዳዮች ምንም አይነት ምላሽ/reply መቀበል አይችሉም",
        });
      }

      const findForwardCase = await ForwardCase.findOne({ case_id: case_id });

      if (!findForwardCase) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Since the case has not been forwarded yet, replies cannot be sent for this case",
          Message_am: "ጉዳዩ እስካሁን ስላልተላለፈ ለዚህ ጉዳይ ምላሽ መላክ አይቻልም",
        });
      }

      const findReplyingPersonInForward = findForwardCase?.path?.filter(
        (item) => item?.to?.toString() === requesterId?.toString()
      );

      if (findReplyingPersonInForward?.length === 0) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "You cannot reply to this case as it has not been forwarded to you",
          Message_am: "ይህ ጉዳይ ለእርስዎ አልተላከም ፤ ስለዚህ ለዚህ ጉዳይ መልስ መስጠት አይችሉም",
        });
      }

      if (findReplyingPersonInForward?.length > 0) {
        const findNormal = findReplyingPersonInForward?.find(
          (item) => item?.cc === "no"
        );

        if (!findNormal) {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en:
              "You cannot reply to this case as it was only CC'd to you, not directly forwarded",
            Message_am:
              "ይህንን ጉዳይ በቀጥታ የተላለፈ/የተላክ ሳይሆን ለእርስዎ CC የተደረገ ብቻ ስለሆነ መልስ መስጠት አይችሉም",
          });
        }
      }

      const findReplyCase = await ReplyCase.findOne({ case_id: case_id });

      if (findReplyCase) {
        const findNormalForward = findReplyingPersonInForward?.find(
          (item) => item?.cc === "no"
        );
        const findSendingUserInReply = findReplyCase?.path?.filter(
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
                "You cannot reply to this case as it was only CC'd to you",
              Message_am: "ለእርስዎ CC ብቻ ስለነበር የተደረገው ለዚህ ጉዳይ መልስ መስጠት አይችሉም",
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
            Message_en: "You can not reply or CC a case to yourself",
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

        const findCCForwardTo = findForwardCase?.path?.filter(
          (item) => item?.to?.toString() === to?.toString()
        );

        if (cc === "no" && findCCForwardTo?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: `You cannot reply to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} as this case was not forwarded to him/her`,
            Message_am: `ይህ ጉዳይ ወደ እሱ/ሷ ስላልተላከ ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ምላሽ መስጠት አይችሉም`,
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
                "Invalid case letter format please try again. Only accepts '.pdf'.",
              Message_am: "ልክ ያልሆነ የጉዳይ አባሪ እባክዎ እንደገና ይሞክሩ። '.pdf' ብቻ ይቀበላል።",
            });
          }

          const bytes = await readFile(attachment?.filepath);
          const caseReplyAttachmentBuffer = Buffer.from(bytes);
          const uniqueSuffix =
            Date.now() + "-" + Math.round(Math.random() * 1e9);

          const path = join(
            "./",
            "Media",
            "CaseReplyFiles",
            uniqueSuffix + "-" + attachment?.originalFilename
          );

          attachmentName = uniqueSuffix + "-" + attachment?.originalFilename;

          await writeFile(path, caseReplyAttachmentBuffer);
        }

        if (!findReplyCase) {
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

          const createReply = await ReplyCase.create({
            case_id: case_id,
            forward_case_id: findForwardCase?._id,
            path,
          });

          const updateHistory = [
            {
              updatedByOfficeUser: requesterId,
              action: "create",
            },
          ];

          try {
            await ReplyCaseHistory.create({
              reply_case_id: createReply?._id,
              updateHistory,
              history: createReply?.toObject(),
            });
          } catch (error) {
            console.log(
              `Reply history for case with case number (${findCase?.case_number}) is not created`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for case with case number ${findCase?.case_number} is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Case with case number ${findCase?.case_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `የጉዳይ ቁጥር ${findCase?.case_number} ባለው ጉዳይ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
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
            io.to(user?.socketID).emit("case_reply_notification", {
              Message_en: `Replies for case with case number ${findCase?.case_number} is sent to you.`,
              Message_am: `የጉዳይ ቁጥር ${findCase?.case_number} ላለው ጉዳይ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.CREATED).json({
            Message_en: `Reply for case with case number ${findCase?.case_number} is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለጉዳይ ቁጥር ${findCase?.case_number} ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
          });
        }

        if (findReplyCase) {
          const updateReply = await ReplyCase.findOneAndUpdate(
            { _id: findReplyCase?._id },
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
              Message_en: `The attempt to reply to case number ${findCase?.case_number} for ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} was unsuccessful.`,
              Message_am: `የጉዳይ ቁጥር ${findCase?.case_number} ለ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}ን ለመመለስ የተደረገው ሙከራ አልተሳካም።`,
            });
          }

          try {
            await ReplyCaseHistory.findOneAndUpdate(
              { reply_case_id: findReplyCase?._id },
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
              `Reply history of case with case number ${findCase?.case_number} is not updated successfully`
            );
          }

          const notificationMessage = {
            Message_en:
              cc === "no"
                ? `Replies for case with case number ${findCase?.case_number} is sent to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`
                : `Case with case number ${findCase?.case_number} is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am:
              cc === "no"
                ? `የጉዳይ ቁጥር ${findCase?.case_number} ባለው ጉዳይ ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ መልስ ተልኮሎታል።`
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
            io.to(user?.socketID).emit("case_reply_notification", {
              Message_en: `Replies for case with case number ${findCase?.case_number} is sent to you.`,
              Message_am: `የጉዳይ ቁጥር ${findCase?.case_number} ላለው ጉዳይ ወደ እርስዎ መልስ ተልኳል።`,
            });
          }

          return res.status(StatusCodes.OK).json({
            Message_en: `Reply for case with case number ${findCase?.case_number} is successfully sent to ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname}.`,
            Message_am: `ለጉዳይ ቁጥር ${findCase?.case_number} ምላሽ በተሳካ ሁኔታ ወደ ${findAcceptorUser?.firstname} ${findAcceptorUser?.middlename} ${findAcceptorUser?.lastname} ተልኳል።`,
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

const getRepliedCase = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYCASE_API;
    const actualAPIKey = req?.headers?.get_rplycase_api;
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

      const repliedCases = await ReplyCase.find({
        "path.to": requesterId,
        "path.cc": "no",
      });

      if (!repliedCases) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied cases",
          Message_am: "ምንም መልሶች የተሰጡባቸው ጉዳዮች የሎዎትም",
        });
      }

      for (const replyPath of repliedCases) {
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

      const caseIds = replyList?.map((fwdCase) => fwdCase?.case_id);

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
          Message_en: "You have no replied cases",
          Message_am: "ምንም መልሶች የተሰጡባቸው ጉዳዮች የሎዎትም",
        });
      }

      const lstOfRepliedCases = [];

      for (const items of findCases) {
        const findReplyCase = await ReplyCase.findOne({
          case_id: items?._id,
          "path.from_office_user": requesterId,
        });

        let caseFind = "no";

        if (findReplyCase) {
          caseFind = "yes";
        }

        const updatedItem = { ...items.toObject(), caseReplied: caseFind };

        lstOfRepliedCases.push(updatedItem);
      }

      return res.status(StatusCodes.OK).json({
        cases: lstOfRepliedCases,
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

const getRepliedCaseCC = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYCASECC_API;
    const actualAPIKey = req?.headers?.get_rplycasecc_api;
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

      const repliedCases = await ReplyCase.find({
        "path.to": requesterId,
        "path.cc": "yes",
      });

      if (!repliedCases) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "You have no replied cases",
          Message_am: "ምንም መልሶች የተሰጡባቸው ጉዳዮች የሎዎትም",
        });
      }

      for (const replyPath of repliedCases) {
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

      const caseIds = replyList?.map((fwdCase) => fwdCase?.case_id);

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
          Message_en: "You have no replied cases",
          Message_am: "ምንም መልሶች የተሰጡባቸው ጉዳዮች የሎዎትም",
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

const getReplyPath = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RPLYPATH_API;
    const actualAPIKey = req?.headers?.get_rplypath_api;
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
      const findReplyCase = await ReplyCase.findOne({ case_id: case_id })
        .populate({
          path: "path.from_office_user",
          select: "_id firstname middlename lastname position username level",
        })
        .populate({
          path: "path.to",
          select: "_id firstname middlename lastname position username level",
        });

      if (!findCase || !findReplyCase) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The case and its replies were not found",
          Message_am: "ጉዳዩ እና ምላሾቹ አልተገኙም",
        });
      }

      const replyCases = findReplyCase?.path;

      return res.status(StatusCodes.OK).json({
        repliedDocs: replyCases,
        replyId: findReplyCase?._id,
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

const printReplyCase = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRINTCASERPLY_API;
    const actualAPIKey = req?.headers?.get_printcaserply_api;
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

      const findReplyCase = await ReplyCase.findOne({ _id: id });

      const findCase = await Case.findOne({ _id: findReplyCase?.case_id });

      if (!findCase || !findReplyCase) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "The case and its replies were not found",
          Message_am: "ጉዳዩ እና ምላሾቹ አልተገኙም",
        });
      }

      const reply_forward_id = req?.params?.reply_forward_id;

      if (!reply_forward_id || !mongoose.isValidObjectId(reply_forward_id)) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en: "Please specify the exact reply path you want to print",
          Message_am: "እባክዎ ማተም የሚፈልጉትን ትክክለኛ የምላሽ መንገድ ይግለጹ",
        });
      }

      const replyToPrint = findReplyCase?.path?.find(
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
          Message_en: `The person who replied to this case (Case Number: ${findCase?.case_number}) is not found among the office administrators.`,
          Message_am: `ለዚህ ጉዳይ ምላሽ የሰጡት ሰው (የጉዳይ ቁጥር፡ ${findCase?.case_number}) በቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const findRecieverUser = await OfficeUser.findOne({
        _id: replyToPrint?.to,
      });

      if (!findRecieverUser) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `The person who received the answer to this case (Case Number: ${findCase?.case_number}) is not found among the office administrators.`,
          Message_am: `ይህንን ጉዳይ መልስ የተቀበለው ሰው (የጉዳይ ቁጥር፡ ${findCase?.case_number}) ከቢሮ አስተዳዳሪዎች መካከል አልተገኘም።`,
        });
      }

      const sent_date = caseSubDate(replyToPrint?.replied_date);
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
        "ReplyCasePrint",
        uniqueSuffix + "-replycase.pdf"
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
        await appendReplyCasePrint(inputPath, text, outputPath);

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
  replyCase,
  getRepliedCase,
  getRepliedCaseCC,
  getReplyPath,
  printReplyCase,
};
