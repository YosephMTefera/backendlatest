const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const Notification = require("../../model/Notifications/Notification");
const InternalMemo = require("../../model/InternalMemo/InternalMemo");
const InternalMemoHistory = require("../../model/InternalMemo/InternalMemoHistory");
const ForwardInternalMemo = require("../../model/ForwardInternalMemo/ForwardInternalMemo");
const ForwardInternalMemoHistory = require("../../model/ForwardInternalMemo/ForwardInternalMemoHistory");

const fs = require("fs");
const { join } = require("path");
const mongoose = require("mongoose");
const formidable = require("formidable");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile, unlink } = require("fs/promises");
const {
  previewInternalMemoResponse,
} = require("../../middleware/previewInternalMemo");
const {
  finalInternalMemoResponse,
} = require("../../middleware/finalInternalMemo");

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

const createInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CRINTMEMO_API;
    const actualAPIKey = req?.headers?.get_crintmemo_api;
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

      const form = new formidable.IncomingForm();

      form.parse(req, async (err, fields, files) => {
        if (err) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Error getting data",
            Message_am: "ዳታውን የማግኘት ስህተት",
          });
        }

        const to_whom = fields?.to_whom?.[0];
        let to_whom_col = fields?.to_whom_col?.[0] || 1;
        const subject = fields?.subject?.[0];
        const body = fields?.body?.[0];
        const internal_cc = fields?.internal_cc?.[0];
        let internal_cc_col = fields?.internal_cc_col?.[0] || 1;
        const verified_by = fields?.verified_by?.[0];
        const main_letter_attachment = files?.main_letter_attachment?.[0];

        if (!to_whom || to_whom?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Please specify to whom the internal memo letter is written",
            Message_am: "የዉስጥ ማስታወሻዉ ደብዳቤዉ ለማን እንደተጻፈ እባክዎ ይግለጹ",
          });
        }

        if (!subject) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please specify the subject of the letter",
            Message_am: "እባክዎ የደብዳቤውን ርዕሰ ጉዳይ ይግለጹ",
          });
        }

        if (!body) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please provide the body of the letter",
            Message_am: "እባክዎ የደብዳቤውን ሃተታ ያስገቡ",
          });
        }

        if (to_whom_col) {
          to_whom_col = parseInt(to_whom_col);
          if (isNaN(to_whom_col)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }

          if (to_whom_col !== 1 && to_whom_col !== 2) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }
        }

        if (internal_cc_col) {
          internal_cc_col = parseInt(internal_cc_col);
          if (isNaN(internal_cc_col)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the cc receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ግልባጭ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }

          if (internal_cc_col !== 1 && internal_cc_col !== 2) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the cc receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ግልባጭ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }
        }

        if (!verified_by || !mongoose.isValidObjectId(verified_by)) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please specify the approval of this letter",
            Message_am: "እባክዎ ይህ ደቢዳቤ በማን ስም ወጪ እንደሚሆን ይግለጹ",
          });
        }

        const findVerifiedBy = await OfficeUser.findOne({ _id: verified_by });

        if (!findVerifiedBy) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en:
              "The person to approve this letter is not found. Please only select from the existing users.",
            Message_am:
              "ይህንን ደብዳቤ የሚያፀድቀው ተጠቃሚ አልተገኘም። እባክዎ በዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ዉስጥ ብቻ ይምረጡ።",
          });
        }

        if (findVerifiedBy?.status === "inactive") {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en: `The person to approve this letter is currently inactive. Please select an active user to approve this letter. (${
              findVerifiedBy?.firstname +
              " " +
              findVerifiedBy?.middlename +
              " " +
              findVerifiedBy?.lastname
            })`,
            Message_am: `ይህንን ደብዳቤ የሚያፀድቀው ሰው በአሁኑ ጊዜ ኢን-አክቲቭ ነው። እባክዎ ይህን ደብዳቤ ለማጽደቅ ንቁ ተጠቃሚ ይምረጡ። (${
              findVerifiedBy?.firstname +
              " " +
              findVerifiedBy?.middlename +
              " " +
              findVerifiedBy?.lastname
            })`,
          });
        }

        let forwardToWhomArray = [];

        if (to_whom?.length > 0) {
          forwardToWhomArray = Array.isArray(to_whom)
            ? to_whom
            : JSON.parse(to_whom);

          if (forwardToWhomArray?.length === 0) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "Please specify to whom the letter is written",
              Message_am: "ደብዳቤው ለማን እንደተጻፈ እባክዎ ይግለጹ",
            });
          }

          const uniqueForwardToWhomMap = new Map();
          forwardToWhomArray.forEach((item) => {
            uniqueForwardToWhomMap.set(item.internal_office.toString(), item);
          });
          const uniqueForwardToWhomArray = Array.from(
            uniqueForwardToWhomMap.values()
          );

          for (const singlePath of uniqueForwardToWhomArray) {
            if (
              !singlePath?.internal_office ||
              !mongoose.isValidObjectId(singlePath?.internal_office)
            ) {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en: `Please specify the receiver of this letter`,
                Message_am: `እባክዎ የዚህን ደብዳቤ ተቀባይ ይጥቀሱ`,
              });
            }

            const findToWhomUsers = await OfficeUser.findOne({
              _id: singlePath?.internal_office,
            });

            if (!findToWhomUsers) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The receiver of the letter is not found. Please only select from existing users.`,
                Message_am: `የደብዳቤው ተቀባይ ተጠቃሚ አልተገኘም። እባክዎ ዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ብቻ ይምረጡ።`,
              });
            }

            if (findToWhomUsers?.status === "inactive") {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en: `The receiver of this letter is currently inactive and cannot receive the letter. (${
                  findToWhomUsers?.firstname +
                  " " +
                  findToWhomUsers?.middlename +
                  " " +
                  findToWhomUsers?.lastname
                })`,
                Message_am: `የዚህ ደብዳቤ ተቀባይ በአሁኑ ጊዜ ኢን-አክቲቭ ነው እና ደብዳቤውን መቀበል አይችልም። (${
                  findToWhomUsers?.firstname +
                  " " +
                  findToWhomUsers?.middlename +
                  " " +
                  findToWhomUsers?.lastname
                })`,
              });
            }
          }

          forwardToWhomArray = uniqueForwardToWhomArray;
        }

        let forwardIntCC = [];

        if (internal_cc && internal_cc?.length > 0) {
          forwardIntCC = Array.isArray(internal_cc)
            ? internal_cc
            : JSON.parse(internal_cc);

          if (forwardIntCC?.length > 0) {
            const uniqueForwardIntCCMap = new Map();
            forwardIntCC.forEach((item) => {
              uniqueForwardIntCCMap.set(item.internal_office.toString(), item);
            });
            const uniqueForwardIntCCArray = Array.from(
              uniqueForwardIntCCMap.values()
            );

            for (const singleFrwdInc of uniqueForwardIntCCArray) {
              if (
                !singleFrwdInc?.internal_office ||
                !mongoose.isValidObjectId(singleFrwdInc?.internal_office)
              ) {
                return res.status(StatusCodes.BAD_REQUEST).json({
                  Message_en: `Please provide the list persons(with in the organization) to receive the CC of this letter`,
                  Message_am: `እባክዎ የዚህን ደብዳቤ ግልባጭ የሚቀበሉ (በድርጅቱ ውስጥ ያሉ) ሰዎች ዝርዝር ያቅርቡ`,
                });
              }

              const findInternalCCUsers = await OfficeUser.findOne({
                _id: singleFrwdInc?.internal_office,
              });

              if (!findInternalCCUsers) {
                return res.status(StatusCodes.NOT_FOUND).json({
                  Message_en:
                    "The user to receive CC of the letter is not found",
                  Message_am: "የዚህን ደብዳቤ ግልባጭ የሚቀበለው ተጠቃሚ አልተገኘም",
                });
              }

              if (findInternalCCUsers?.status === "inactive") {
                return res.status(StatusCodes.FORBIDDEN).json({
                  Message_en: `The user to receive the CC of the letter is currently inactive. (${
                    findInternalCCUsers?.firstname +
                    " " +
                    findInternalCCUsers?.middlename +
                    " " +
                    findInternalCCUsers?.lastname
                  })`,
                  Message_am: `የዚህን ደብዳቤ ግልባጭ የሚቀበለው ተጠቃሚ በአሁኑ ጊዜ አክቲቭ አይደለም። (${
                    findInternalCCUsers?.firstname +
                    " " +
                    findInternalCCUsers?.middlename +
                    " " +
                    findInternalCCUsers?.lastname
                  })`,
                });
              }
            }
            forwardIntCC = uniqueForwardIntCCArray;
          }
        }

        let attachmentName = "";
        if (main_letter_attachment) {
          if (
            typeof main_letter_attachment === "object" &&
            (main_letter_attachment?.mimetype === "application/pdf" ||
              main_letter_attachment?.mimetype === "application/PDF")
          ) {
            if (main_letter_attachment?.size > 10 * 1024 * 1024) {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en:
                  "The attachment size is too large. Please insert a file less than 10MB.",
                Message_am:
                  "የአባሪው መጠኑ በጣም ትልቅ ነው። እባክዎ ከ10ሜባ በታች የሆነ ፋይል ያስገቡ።",
              });
            }
          } else {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Invalid attachment format please try again. Only accepts '.pdf'.",
              Message_am:
                "ልክ ያልሆነ የደብዳቤ አባሪ ቅርጸት እባክዎ እንደገና ይሞክሩ። '.pdf' ብቻ ይቀበላል።",
            });
          }

          const bytes = await readFile(main_letter_attachment?.filepath);
          const letterReplyAttachmentBuffer = Buffer.from(bytes);
          const uniqueSuffix =
            Date.now() + "-" + Math.round(Math.random() * 1e9);

          const path = join(
            "./",
            "Media",
            "InternalMemoAttachmentFiles",
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename
          );

          attachmentName =
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename;

          await writeFile(path, letterReplyAttachmentBuffer);
        }

        const createIntMemo = await InternalMemo.create({
          to_whom: forwardToWhomArray?.map((item) => ({
            internal_office: item?.internal_office,
          })),
          to_whom_col: to_whom_col,
          subject: subject,
          body: body,
          internal_cc: forwardIntCC?.map((item) => ({
            internal_office: item?.internal_office,
          })),
          verified_by: verified_by,
          internal_cc_col: internal_cc_col,
          main_letter_attachment: attachmentName,
          createdBy: requesterId,
        });

        const updateHistory = [
          {
            updatedByOfficeUser: requesterId,
            action: "create",
          },
        ];

        try {
          await InternalMemoHistory.create({
            internal_memo_id: createIntMemo?._id,
            updateHistory,
            history: createIntMemo?.toObject(),
          });
        } catch (error) {
          console.log(
            `Internal memo history with this ID ${createIntMemo?._id} is not created`
          );
        }

        return res.status(StatusCodes.CREATED).json({
          Message_en: `The internal memo is created with subject ${createIntMemo?.subject} successfully.`,
          Message_am: `ይህ ርዕስ ${createIntMemo?.subject} ያለዉ የዉስጥ ማስታወሻ በተሳካ ሁኔታ ተፈጥሯል።`,
        });
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

const getInternalMemos = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INTMEMOS_API;
    const actualAPIKey = req?.headers?.get_intmemos_api;
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

const getInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INTMEMO_API;
    const actualAPIKey = req?.headers?.get_intmemo_api;
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

      const findInternalMemo = await InternalMemo.findOne({ _id: id })
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
          Message_en: "Internal memo is not found",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤው አልተገኘም",
        });
      }

      return res.status(StatusCodes.OK).json(findInternalMemo);
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

const updInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_UPDINTMEMO_API;
    const actualAPIKey = req?.headers?.get_updintmemo_api;
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

      const findInternalMemo = await InternalMemo.findOne({ _id: id });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal memo is not found",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalMemo?.status !== "pending") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated. So, this letter cannot be updated.",
          Message_am: "ደብዳቤው አስቀድሞ ተፈጥሯል። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      const findForwardInternalMemo = await ForwardInternalMemo.findOne({
        internal_memo_id: id,
      });

      const findForwardedPerson = findForwardInternalMemo?.path?.find(
        (item) =>
          item?.to?.toString() === requesterId?.toString() && item?.cc === "no"
      );

      if (
        findInternalMemo?.createdBy?.toString() !== requesterId?.toString() &&
        !findForwardedPerson
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "This letter is neither created by you nor forwarded to you, so you cannot edit the letter. If you are CC'd on this letter, you cannot edit this letter.",
          Message_am:
            "ይህ ደብዳቤ በእርስዎ የተፈጠረ ወይም ወደ እርስዎ የተላከ አይደለም፣ ስለዚህ ደብዳቤውን ማስተካከል አይችሉም። ይህ ደብዳቤ ግልባጭ ከተደረገሎት ማደስ (edit) ማድረግ አይችሉም።",
        });
      }

      const form = new formidable.IncomingForm();

      form.parse(req, async (err, fields, files) => {
        if (err) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Error getting data",
            Message_am: "ዳታውን የማግኘት ስህተት",
          });
        }

        const subject = fields?.subject?.[0];
        const body = fields?.body?.[0];
        const main_letter_attachment = files?.main_letter_attachment?.[0];
        const detachFile = fields?.detach_file?.[0];
        let to_whom_col = fields?.to_whom_col?.[0];
        const output_by = fields?.output_by?.[0];
        let internal_cc_col = fields?.internal_cc_col?.[0];

        const updatedFields = {};

        if (subject) {
          updatedFields.subject = subject;
        }
        if (body) {
          updatedFields.body = body;
        }
        if (to_whom_col) {
          to_whom_col = parseInt(to_whom_col);
          if (isNaN(to_whom_col)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }

          if (to_whom_col !== 1 && to_whom_col !== 2) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }

          updatedFields.to_whom_col = to_whom_col;
        }

        if (internal_cc_col) {
          internal_cc_col = parseInt(internal_cc_col);
          if (isNaN(internal_cc_col)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the cc receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ግልባጭ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }

          if (internal_cc_col !== 1 && internal_cc_col !== 2) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Please insert a valid column for the cc receivers' of this letter. (Please choose 1 or 2 as a valid column)",
              Message_am:
                "እባክዎ ለዚህ ደብዳቤ ግልባጭ ተቀባዮች ዝርዝር ትክክለኛ ኮለን ያስገቡ። (1 ወይም 2 ብለው ኮለን ይምረጡ)",
            });
          }

          updatedFields.internal_cc_col = internal_cc_col;
        }
        if (detachFile) {
          if (detachFile === "detachAttachment") {
            if (findInternalMemo?.main_letter_attachment) {
              const detPath = join(
                "./",
                "Media",
                "InternalMemoAttachmentFiles",
                findInternalMemo?.main_letter_attachment
              );

              try {
                await unlink(detPath);
                findInternalMemo.main_letter_attachment = "";
                await findInternalMemo.save();
              } catch (error) {
                console.log(
                  `Internal memo with ID ${findInternalMemo?._id} attachment is not found`
                );
              }
            }
          }
        }

        if (output_by) {
          if (
            findInternalMemo?.createdBy?.toString() === requesterId?.toString()
          ) {
            if (!output_by || !mongoose.isValidObjectId(output_by)) {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en: "Please specify the approval of this letter",
                Message_am: "እባክዎ ይህ ደቢዳቤ በማን ስም ወጪ እንደሚሆን ይግለጹ",
              });
            }

            const findOutputBy = await OfficeUser.findOne({ _id: output_by });

            if (!findOutputBy) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en:
                  "The person to approve this letter is not found. Please only select from the existing users.",
                Message_am:
                  "ይህንን ደብዳቤ የሚያፀድቀው ተጠቃሚ አልተገኘም። እባክዎ በዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ዉስጥ ብቻ ይምረጡ።",
              });
            }

            if (findOutputBy?.status === "inactive") {
              return res.status(StatusCodes.FORBIDDEN).json({
                Message_en: `The person to approve this letter is currently inactive. Please select an active user to approve this letter. (${
                  findOutputBy?.firstname +
                  " " +
                  findOutputBy?.middlename +
                  " " +
                  findOutputBy?.lastname
                })`,
                Message_am: `ይህንን ደብዳቤ የሚያፀድቀው ሰው በአሁኑ ጊዜ ኢን-አክቲቭ ነው። እባክዎ ይህን ደብዳቤ ለማጽደቅ ንቁ ተጠቃሚ ይምረጡ። (${
                  findOutputBy?.firstname +
                  " " +
                  findOutputBy?.middlename +
                  " " +
                  findOutputBy?.lastname
                })`,
              });
            }

            updatedFields.verified_by = output_by;
          }
        }

        const updateArrayField = async (array, items, fields, fieldName) => {
          let updatedArray = [...array];

          for (const item of items) {
            if (item?.action === "add") {
              if (fieldName === "internal_cc") {
                if (
                  !item?.value?.internal_office ||
                  !mongoose.isValidObjectId(item?.value?.internal_office)
                ) {
                  throw {
                    status: StatusCodes.BAD_REQUEST,
                    message: {
                      Message_en:
                        "The account of person to be added as a recipient (from the office) of CC of this letter is not found",
                      Message_am:
                        "አዲስ የሚጨመርው የዚህ ደብዳቤ CC (Internal CC) ተቀባይ ሰው መለያ አልተገኘም።",
                    },
                  };
                }

                const findInternalCCUsers = await OfficeUser.findOne({
                  _id: item?.value?.internal_office,
                });

                if (!findInternalCCUsers) {
                  throw {
                    status: StatusCodes.NOT_FOUND,
                    message: {
                      Message_en:
                        "The user to receive CC of the letter is not found",
                      Message_am: "የዚህን ደብዳቤ ግልባጭ የሚቀበለው ተጠቃሚ አልተገኘም",
                    },
                  };
                }

                if (findInternalCCUsers?.status === "inactive") {
                  throw {
                    status: StatusCodes.FORBIDDEN,
                    message: {
                      Message_en: `The user to receive the CC of the letter is currently inactive. (${
                        findInternalCCUsers?.firstname +
                        " " +
                        findInternalCCUsers?.middlename +
                        " " +
                        findInternalCCUsers?.lastname
                      })`,
                      Message_am: `የዚህን ደብዳቤ ግልባጭ የሚቀበለው ተጠቃሚ በአሁኑ ጊዜ አክቲቭ አይደለም። (${
                        findInternalCCUsers?.firstname +
                        " " +
                        findInternalCCUsers?.middlename +
                        " " +
                        findInternalCCUsers?.lastname
                      })`,
                    },
                  };
                }

                if (
                  findInternalMemo?.internal_cc &&
                  findInternalMemo?.internal_cc?.length > 0
                ) {
                  for (const intUser of findInternalMemo?.internal_cc) {
                    if (
                      item?.value?.internal_office?.toString() ===
                      intUser?.internal_office?.toString()
                    ) {
                      throw {
                        status: StatusCodes.BAD_REQUEST,
                        message: {
                          Message_en: `The user to receive the CC of this letter is already mentioned on the receivers list. (${
                            findInternalCCUsers?.firstname +
                            " " +
                            findInternalCCUsers?.middlename +
                            " " +
                            findInternalCCUsers?.lastname
                          })`,
                          Message_am: `የዚህን ደብዳቤ ግልባጭ ለመቀበል ተጠቃሚው አስቀድሞ በተቀባዮች ዝርዝር ዉስጥ ተጠቅሷል። (${
                            findInternalCCUsers?.firstname +
                            " " +
                            findInternalCCUsers?.middlename +
                            " " +
                            findInternalCCUsers?.lastname
                          })`,
                        },
                      };
                    }
                  }
                }
              }

              if (fieldName === "to_whom") {
                if (
                  !item?.value?.internal_office ||
                  !mongoose.isValidObjectId(item?.value?.internal_office)
                ) {
                  throw {
                    status: StatusCodes.BAD_REQUEST,
                    message: {
                      Message_en:
                        "The account of person to be added as a recipient of this letter is not found. Please only choose from existing users.",
                      Message_am:
                        "አዲስ የሚጨመርው የዚህ ደብዳቤ ተቀባይ መለያዉ አልተገኘም። እባክዎ ካሉ በዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ብቻ ይምረጡ።",
                    },
                  };
                }

                const findToWhomLstUsers = await OfficeUser.findOne({
                  _id: item?.value?.internal_office,
                });

                if (!findToWhomLstUsers) {
                  throw {
                    status: StatusCodes.BAD_REQUEST,
                    message: {
                      Message_en:
                        "The account of person to be added as a recipient of this letter is not found. Please only choose from existing users.",
                      Message_am:
                        "አዲስ የሚጨመርው የዚህ ደብዳቤ ተቀባይ መለያዉ አልተገኘም። እባክዎ በዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ብቻ ይምረጡ።",
                    },
                  };
                }

                if (findToWhomLstUsers?.status === "inactive") {
                  throw {
                    status: StatusCodes.FORBIDDEN,
                    message: {
                      Message_en: `The user to receive the letter is currently inactive. (${
                        findToWhomLstUsers?.firstname +
                        " " +
                        findToWhomLstUsers?.middlename +
                        " " +
                        findToWhomLstUsers?.lastname
                      })`,
                      Message_am: `ይህንን ደብዳቤ የሚቀበለው ተጠቃሚ በአሁኑ ጊዜ አክቲቭ አይደለም። (${
                        findToWhomLstUsers?.firstname +
                        " " +
                        findToWhomLstUsers?.middlename +
                        " " +
                        findToWhomLstUsers?.lastname
                      })`,
                    },
                  };
                }

                if (
                  findInternalMemo?.to_whom &&
                  findInternalMemo?.to_whom?.length > 0
                ) {
                  for (const toWhom of findInternalMemo?.to_whom) {
                    if (
                      item?.value?.internal_office?.toString() ===
                      toWhom?.internal_office?.toString()
                    ) {
                      throw {
                        status: StatusCodes.BAD_REQUEST,
                        message: {
                          Message_en: `The user to receive this letter is already mentioned on the receivers' list. (${
                            findToWhomLstUsers?.firstname +
                            " " +
                            findToWhomLstUsers?.middlename +
                            " " +
                            findToWhomLstUsers?.lastname
                          })`,
                          Message_am: `ይህንን ደብዳቤ ለመቀበል ተጠቃሚው አስቀድሞ በተቀባዮች ዝርዝር ዉስጥ ተጠቅሷል። (${
                            findToWhomLstUsers?.firstname +
                            " " +
                            findToWhomLstUsers?.middlename +
                            " " +
                            findToWhomLstUsers?.lastname
                          })`,
                        },
                      };
                    }
                  }
                }
              }
              updatedArray.push(item?.value);
            } else if (item?.action === "remove") {
              if (!item?.value?._id) {
                throw {
                  status: StatusCodes.NOT_ACCEPTABLE,
                  message: {
                    Message_en: "Invalid request",
                    Message_am: "ልክ ያልሆነ ጥያቄ",
                  },
                };
              }
              updatedArray = updatedArray.filter(
                (element) =>
                  element._id?.toString() !== item?.value?._id?.toString()
              );
            } else if (item?.action === "update") {
              const index = updatedArray.findIndex(
                (element) =>
                  element._id?.toString() === item?.value?._id?.toString()
              );
              if (index !== -1) {
                if (fieldName === "internal_cc") {
                  if (
                    !item?.value?.internal_office ||
                    !mongoose.isValidObjectId(item?.value?.internal_office)
                  ) {
                    throw {
                      status: StatusCodes.BAD_REQUEST,
                      message: {
                        Message_en:
                          "The account of person to be updated as a recipient (from the office) of CC of this letter is not found",
                        Message_am:
                          "አዲስ የሚዘመነው የዚህ ደብዳቤ CC (Internal CC) ተቀባይ ሰው መለያ አልተገኘም።",
                      },
                    };
                  }

                  const findInternalCCUsers = await OfficeUser.findOne({
                    _id: item?.value?.internal_office,
                  });

                  if (!findInternalCCUsers) {
                    throw {
                      status: StatusCodes.NOT_FOUND,
                      message: {
                        Message_en:
                          "The user to receive CC of the letter is not found",
                        Message_am: "የዚህን ደብዳቤ ግልባጭ የሚቀበለው ተጠቃሚ አልተገኘም",
                      },
                    };
                  }

                  if (findInternalCCUsers?.status === "inactive") {
                    throw {
                      status: StatusCodes.FORBIDDEN,
                      message: {
                        Message_en: `The user to receive the CC of the letter is currently inactive. (${
                          findInternalCCUsers?.firstname +
                          " " +
                          findInternalCCUsers?.middlename +
                          " " +
                          findInternalCCUsers?.lastname
                        })`,
                        Message_am: `የዚህን ደብዳቤ ግልባጭ የሚቀበለው ተጠቃሚ በአሁኑ ጊዜ አክቲቭ አይደለም። (${
                          findInternalCCUsers?.firstname +
                          " " +
                          findInternalCCUsers?.middlename +
                          " " +
                          findInternalCCUsers?.lastname
                        })`,
                      },
                    };
                  }
                }

                if (fieldName === "to_whom") {
                  if (
                    !item?.value?.internal_office ||
                    !mongoose.isValidObjectId(item?.value?.internal_office)
                  ) {
                    throw {
                      status: StatusCodes.BAD_REQUEST,
                      message: {
                        Message_en:
                          "The account of person to be added as a recipient of this letter is not found. Please only choose from existing users.",
                        Message_am:
                          "አዲስ የሚጨመርው የዚህ ደብዳቤ ተቀባይ መለያዉ አልተገኘም። እባክዎ ካሉ በዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ብቻ ይምረጡ።",
                      },
                    };
                  }

                  const findToWhomLstUsers = await OfficeUser.findOne({
                    _id: item?.value?.internal_office,
                  });

                  if (!findToWhomLstUsers) {
                    throw {
                      status: StatusCodes.BAD_REQUEST,
                      message: {
                        Message_en:
                          "The account of person to be added as a recipient of this letter is not found. Please only choose from existing users.",
                        Message_am:
                          "አዲስ የሚጨመርው የዚህ ደብዳቤ ተቀባይ መለያዉ አልተገኘም። እባክዎ በዝርዝር ዉስጥ ካሉ ተጠቃሚዎች ብቻ ይምረጡ።",
                      },
                    };
                  }

                  if (findToWhomLstUsers?.status === "inactive") {
                    throw {
                      status: StatusCodes.FORBIDDEN,
                      message: {
                        Message_en: `The user to receive the letter is currently inactive. (${
                          findToWhomLstUsers?.firstname +
                          " " +
                          findToWhomLstUsers?.middlename +
                          " " +
                          findToWhomLstUsers?.lastname
                        })`,
                        Message_am: `ይህንን ደብዳቤ የሚቀበለው ተጠቃሚ በአሁኑ ጊዜ አክቲቭ አይደለም። (${
                          findToWhomLstUsers?.firstname +
                          " " +
                          findToWhomLstUsers?.middlename +
                          " " +
                          findToWhomLstUsers?.lastname
                        })`,
                      },
                    };
                  }
                }

                updatedArray[index] = item?.value;
              }
            }
          }

          fields[fieldName] = updatedArray;
        };

        try {
          if (fields?.to_whom?.[0]) {
            const toWhomUpdates = JSON.parse(fields?.to_whom?.[0]);
            await updateArrayField(
              findInternalMemo.to_whom,
              toWhomUpdates,
              updatedFields,
              "to_whom"
            );
          }

          if (fields?.internal_cc?.[0]) {
            const internalCcUpdates = JSON.parse(fields?.internal_cc?.[0]);
            await updateArrayField(
              findInternalMemo.internal_cc,
              internalCcUpdates,
              updatedFields,
              "internal_cc"
            );
          }
        } catch (error) {
          return res
            .status(error?.status || StatusCodes.INTERNAL_SERVER_ERROR)
            .json(error?.message);
        }

        if (main_letter_attachment) {
          if (
            typeof main_letter_attachment === "object" &&
            (main_letter_attachment?.mimetype === "application/pdf" ||
              main_letter_attachment?.mimetype === "application/PDF")
          ) {
            if (main_letter_attachment?.size > 10 * 1024 * 1024) {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en:
                  "Letter's attachment file size is too large. Please insert a file less than 10MB",
                Message_am:
                  "የደብዳቤው አባሪ ፋይል መጠን በጣም ትልቅ ነው። እባክህ ከ 10ሜባ በታች የሆነ ፋይል አስገባ",
              });
            }
          } else {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Invalid letter attachment file format please try again. Only accepts '.pdf'",
              Message_am:
                "ልክ ያልሆነ የደብዳቤው አባሪ ፋይል ቅርጸት እባክዎ እንደገና ይሞክሩ። «.pdf»ን ብቻ ይቀበላል",
            });
          }

          const bytes = await readFile(main_letter_attachment?.filepath);
          const mainLetterBuffer = Buffer.from(bytes);
          const uniqueSuffix =
            Date.now() + "-" + Math.round(Math.random() * 1e9);

          const newPath = join(
            "./",
            "Media",
            "InternalMemoAttachmentFiles",
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename
          );

          const mainLetterAttachmentName =
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename;

          await writeFile(newPath, mainLetterBuffer);

          updatedFields.main_letter_attachment = main_letter_attachment
            ? mainLetterAttachmentName
            : findInternalMemo?.main_letter_attachment;

          if (findInternalMemo?.main_letter_attachment) {
            const path = join(
              "./",
              "Media",
              "InternalMemoAttachmentFiles",
              findInternalMemo?.main_letter_attachment
            );

            try {
              await unlink(path);
            } catch (error) {
              console.log(
                `Internal memo with ID ${findInternalMemo?._id} attachment is replaced, but previous attachment is not found`
              );
            }
          }
        }

        const updatedByLists = [];

        if (
          findInternalMemo?.updated_by &&
          Array.isArray(findInternalMemo?.updated_by)
        ) {
          for (const existingUpdated of findInternalMemo?.updated_by) {
            updatedByLists.push({
              updated_date: existingUpdated?.updated_date,
              update_officer: existingUpdated?.update_officer,
            });
          }
        }

        updatedByLists.push({
          updated_date: new Date(),
          update_officer: requesterId,
        });

        updatedFields.updated_by = updatedByLists;

        const newUpdatedInternalMemo = await InternalMemo.findOneAndUpdate(
          { _id: id },
          updatedFields,
          { new: true }
        );

        if (!newUpdatedInternalMemo) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: "Internal memo is not found",
            Message_am: "የዉስጥ ማስታወሻ ደብዳቤው አልተገኘም",
          });
        }

        try {
          await InternalMemoHistory.findOneAndUpdate(
            { internal_memo_id: id },
            {
              $push: {
                updateHistory: {
                  updatedByOfficeUser: requesterId,
                  action: "update",
                },
                history: newUpdatedInternalMemo?.toObject(),
              },
            }
          );
        } catch (error) {
          console.log(
            `Internal memo history with this ID ${findInternalMemo?._id} is not updated successfully`
          );
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `Internal memo with subject ${newUpdatedInternalMemo?.subject} is updated successfully.`,
          Message_am: `ይህ ርዕስ ${newUpdatedInternalMemo?.subject} ያለዉ የዉስጥ ማስታወሻ በተሳካ ሁኔታ ዘምኗል።`,
        });
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

const prvInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRVINTMEMO_API;
    const actualAPIKey = req?.headers?.get_prvintmemo_api;
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

      const findInternalMemo = await InternalMemo.findOne({ _id: id });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal memo is not found",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalMemo?.status !== "pending") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated. So, this letter cannot be updated.",
          Message_am: "ደብዳቤው አስቀድሞ ተፈጥሯል። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      if (findInternalMemo?.to_whom?.length === 0) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the list of recipients to this letter before previewing this letter",
          Message_am:
            "እባክዎ ይህንን ደብዳቤ አስቀድመው ከማየትዎ በፊት የዚህን ደብዳቤ የተቀባዮች ዝርዝር ይግለጹ",
        });
      }

      if (!findInternalMemo?.subject) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the subject of the letter before previewing this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ አስቀድመው ከማየትዎ በፊት የደብዳቤውን ርዕሰ ጉዳይ ይግለጹ",
        });
      }

      if (!findInternalMemo?.body) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the body of the letter before previewing this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ አስቀድመው ከማየትዎ በፊት የደብዳቤውን ሃተታ ያስገቡ",
        });
      }

      let lstOfToWhom = [];
      if (findInternalMemo?.to_whom?.length > 0) {
        for (const singlePerson of findInternalMemo?.to_whom) {
          const findUsers = await OfficeUser.findOne({
            _id: singlePerson?.internal_office,
          });

          if (findUsers) {
            lstOfToWhom?.push({
              internal_office: findUsers?.position,
            });
          }
        }
      }

      let lstOfIntCC = [];
      if (findInternalMemo?.internal_cc?.length > 0) {
        for (const singlePerson of findInternalMemo?.internal_cc) {
          const findUsers = await OfficeUser.findOne({
            _id: singlePerson?.internal_office,
          });

          if (findUsers) {
            lstOfIntCC?.push({
              internal_office: findUsers?.position,
            });
          }
        }
      }

      let attachmentPath = "";

      if (findInternalMemo?.main_letter_attachment) {
        attachmentPath = join(
          "./",
          "Media",
          "InternalMemoAttachmentFiles",
          findInternalMemo?.main_letter_attachment
        );
      }

      const findVerifiedBy = await OfficeUser.findOne({
        _id: findInternalMemo?.verified_by,
      });

      const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);

      const inputPath = join(
        "./",
        "Media",
        "GenerateFiles",
        "Letterhead_AAHDC_Memo.pdf"
      );

      const outputPath = join(
        "./",
        "Media",
        "PreviewInternalMemos",
        uniqueSuffix + "-previewinternalmemos.pdf"
      );

      const text = [
        { toWhom: lstOfToWhom },
        { subject: findInternalMemo?.subject },
        { body: findInternalMemo?.body },
        { toExt: lstOfIntCC },
        { toWhomColNum: findInternalMemo?.to_whom_col },
        { internalCCNum: findInternalMemo?.internal_cc_col },
        { attachmentFile: attachmentPath },
        { senderUser: findVerifiedBy?.position },
      ];

      try {
        await previewInternalMemoResponse(inputPath, text, outputPath);

        const modifiedOutputPath = outputPath
          .replace(/\\/g, "/")
          .replace(/Media\//, "/");

        return res.status(StatusCodes.OK).json(modifiedOutputPath);
      } catch (error) {
        return res.status(StatusCodes.EXPECTATION_FAILED).json({
          Message_en:
            "The system is currently unable to generate the internal memo preview. Please try again.",
          Message_am:
            "ስርዓቱ በአሁኑ ጊዜ ለዉስጥ ማስታወሻ ደብዳቤ ቅድመ እይታ መፍጠር አልቻለም። እባክዎ ዳግም ይሞክሩ።",
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

const vrfyInternalMemo = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRVINTMEMO_API;
    const actualAPIKey = req?.headers?.get_prvintmemo_api;
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

      const findInternalMemo = await InternalMemo.findOne({ _id: id });

      if (!findInternalMemo) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal memo is not found",
          Message_am: "የዉስጥ ማስታወሻ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalMemo?.status !== "pending") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated. So, this letter cannot be updated.",
          Message_am: "ደብዳቤው አስቀድሞ ተፈጥሯል። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      if (findInternalMemo?.to_whom?.length === 0) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the list of recipients to this letter before verifying this letter",
          Message_am:
            "እባክዎ ይህንን ደብዳቤ በስምዎ ወጪ ከማድረግዎ በፊት የዚህን ደብዳቤ የተቀባዮች ዝርዝር ይግለጹ",
        });
      }

      if (!findInternalMemo?.subject) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the subject of the letter before verifying this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ በስምዎ ወጪ ከማድረግዎ በፊት የደብዳቤውን ርዕሰ ጉዳይ ይግለጹ",
        });
      }

      if (!findInternalMemo?.body) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the body of the letter before verifying this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ በስምዎ ወጪ ከማድረግዎ በፊት የደብዳቤውን ሃተታ ያስገቡ",
        });
      }

      if (
        findInternalMemo?.verified_by?.toString() !== requesterId?.toString()
      ) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en: "You are not the one selected to approve this letter.",
          Message_am: "ደብዳቤዉ በእርስዎ ስም ወጪ እንዲያደርጉ ስላልተመረጡ ደብዳቤዉን ወጪ ማድረግ አይችሉም።",
        });
      }

      const checkSignature = join(
        "./",
        "Media",
        "OfficeUserSignatures",
        findRequesterOfficeUser?.signature
      );

      if (!fs.existsSync(checkSignature)) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en:
            "Your signature is not found. Please contact the IT Officers to add a signature to your account profile.",
          Message_am:
            "የእርስዎ ፊርማ አልተገኘም። እባክዎ ከአይቲ(IT) ኦፊሰሮች ጋር በማነጋገር ወደ መለያዎ ፊርማ ያስገቡ።",
        });
      }

      let checkTiter = "";
      if (findRequesterOfficeUser?.titer) {
        checkTiter = join(
          "./",
          "Media",
          "OfficeUserTiters",
          findRequesterOfficeUser?.titer
        );

        if (!fs.existsSync(checkTiter)) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en:
              "Your titer is not found. Please contact the IT Officers to add a titer to your account profile.",
            Message_am:
              "የእርስዎ ቲተር አልተገኘም። እባክዎ ከአይቲ(IT) ኦፊሰሮች ጋር በማነጋገር ወደ መለያዎ ቲተር ያስገቡ።",
          });
        }
      }

      let lstOfToWhom = [];
      if (findInternalMemo?.to_whom?.length > 0) {
        for (const singlePerson of findInternalMemo?.to_whom) {
          const findUsers = await OfficeUser.findOne({
            _id: singlePerson?.internal_office,
          });

          if (findUsers) {
            lstOfToWhom?.push({
              internal_office: findUsers?.position,
            });
          }
        }
      }

      let lstOfIntCC = [];
      if (findInternalMemo?.internal_cc?.length > 0) {
        for (const singlePerson of findInternalMemo?.internal_cc) {
          const findUsers = await OfficeUser.findOne({
            _id: singlePerson?.internal_office,
          });

          if (findUsers) {
            lstOfIntCC?.push({
              internal_office: findUsers?.position,
            });
          }
        }
      }

      const orgStructure = [
        "MainExecutive",
        "DivisionManagers",
        "Directors",
        "TeamLeaders",
        "Professionals",
      ];
      const lstOfIntCCWithLevels = [];
      for (const person of lstOfIntCC) {
        const user = await OfficeUser.findOne({
          position: person?.internal_office,
        });
        lstOfIntCCWithLevels.push({ ...person, level: user?.level });
      }
      lstOfIntCCWithLevels.sort(
        (a, b) =>
          orgStructure.indexOf(a?.level) - orgStructure.indexOf(b?.level)
      );

      let attachmentPath = "";

      if (findInternalMemo?.main_letter_attachment) {
        attachmentPath = join(
          "./",
          "Media",
          "InternalMemoAttachmentFiles",
          findInternalMemo?.main_letter_attachment
        );
      }

      const findVerifiedBy = await OfficeUser.findOne({
        _id: findInternalMemo?.verified_by,
      });

      const io = global?.io;
      const onlineUserList = global?.onlineUserList;
      const verified_date = new Date();
      let verDate = caseSubDate(verified_date);

      const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);

      const inputPath = join(
        "./",
        "Media",
        "GenerateFiles",
        "Letterhead_AAHDC_Memo.pdf"
      );

      const outputPath = join(
        "./",
        "Media",
        "InternalMemoFiles",
        uniqueSuffix + "-internalmemos.pdf"
      );

      const text = [
        { toWhom: lstOfToWhom },
        { subject: findInternalMemo?.subject },
        { body: findInternalMemo?.body },
        { verSign: checkSignature },
        { verTiter: checkTiter },
        { verDate: verDate },
        { senderUser: findVerifiedBy?.position },
        { toExt: lstOfIntCCWithLevels },
        { toWhomColNum: findInternalMemo?.to_whom_col },
        { internalCCNum: findInternalMemo?.internal_cc_col },
        { attachmentFile: attachmentPath },
      ];

      try {
        await finalInternalMemoResponse(inputPath, text, outputPath);

        const modifiedOutputPath = outputPath
          .replace(/\\/g, "/")
          .replace(/Media\//, "/");

        const updatedFields = {};

        updatedFields.main_letter = modifiedOutputPath;
        updatedFields.verified_date = new Date();
        updatedFields.status = "verified";

        const verfyInternalMemo = await InternalMemo.findOneAndUpdate(
          { _id: id },
          updatedFields,
          { new: true }
        );

        if (!verfyInternalMemo) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: `Internal memo is not found`,
            Message_am: `የዉስጥ ማስታወሻ ደብዳቤዉ አልተገኘም`,
          });
        }

        try {
          await InternalMemoHistory.findOneAndUpdate(
            { internal_memo_id: id },
            {
              $push: {
                updateHistory: {
                  updatedByOfficeUser: requesterId,
                  action: "update",
                },
                history: verfyInternalMemo?.toObject(),
              },
            }
          );
        } catch (error) {
          console.log(
            `Internal memo with ID ${findInternalMemo?._id} is not updated successfully`
          );
        }

        for (const toWhom of findInternalMemo?.to_whom) {
          const findForwardInternalMemo = await ForwardInternalMemo.findOne({
            internal_memo_id: id,
          });

          if (findForwardInternalMemo) {
            const forwardInternalMemoToOfficer =
              await ForwardInternalMemo.findOneAndUpdate(
                { internal_memo_id: id },
                {
                  $push: {
                    path: {
                      forwarded_date: new Date(),
                      from_office_user: requesterId,
                      cc: "no",
                      to: toWhom?.internal_office,
                    },
                  },
                },
                { new: true }
              );

            try {
              await ForwardInternalMemoHistory.findOneAndUpdate(
                { forward_internal_memo_id: forwardInternalMemoToOfficer?._id },
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

          if (!findForwardInternalMemo) {
            const path = [
              {
                forwarded_date: new Date(),
                from_office_user: requesterId,
                cc: "no",
                to: toWhom?.internal_office,
              },
            ];

            const newInternalMemoForward = await ForwardInternalMemo.create({
              internal_memo_id: id,
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
                updateHistory,
                history: newInternalMemoForward?.toObject(),
              });
            } catch (error) {
              console.log(
                `Forward history for internal memo with ID ${findInternalMemo?._id} is not created successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en: `An internal memo is forwarded to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am: `የዉስጥ ማስታወሻ ደብዳቤ (ኦፊሻል የሆነ) ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ ተልኳል።`,
          };

          await Notification.create({
            office_user: toWhom?.internal_office,
            notifcation_type: "InternalMemo",
            document_id: findInternalMemo?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(toWhom?.internal_office, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("internal_memo_forward_notification", {
              Message_en: `Internal memo is forwarded to you from ${
                findRequesterOfficeUser?.firstname +
                " " +
                findRequesterOfficeUser?.middlename +
                " " +
                findRequesterOfficeUser?.lastname +
                "-" +
                findRequesterOfficeUser?.position
              }.`,
              Message_am: `የዉስጥ ማስታወሻ ደብዳቤው ወደ እርስዎ ከ${
                findRequesterOfficeUser?.firstname +
                " " +
                findRequesterOfficeUser?.middlename +
                " " +
                findRequesterOfficeUser?.lastname +
                "-" +
                findRequesterOfficeUser?.position
              } ተልኳል።`,
            });
          }
        }

        for (const internalCC of findInternalMemo?.internal_cc) {
          const findForwardInternalMemo = await ForwardInternalMemo.findOne({
            internal_memo_id: id,
          });

          if (findForwardInternalMemo) {
            const forwardInternalMemoToOfficer =
              await ForwardInternalMemo.findOneAndUpdate(
                { internal_memo_id: id },
                {
                  $push: {
                    path: {
                      forwarded_date: new Date(),
                      from_office_user: requesterId,
                      cc: "yes",
                      to: internalCC?.internal_office,
                    },
                  },
                },
                { new: true }
              );

            try {
              await ForwardInternalMemoHistory.findOneAndUpdate(
                { forward_internal_memo_id: forwardInternalMemoToOfficer?._id },
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

          if (!findForwardInternalMemo) {
            const path = [
              {
                forwarded_date: new Date(),
                from_office_user: requesterId,
                cc: "yes",
                to: internalCC?.internal_office,
              },
            ];

            const newInternalMemoForward = await ForwardInternalMemo.create({
              internal_memo_id: id,
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
                updateHistory,
                history: newInternalMemoForward?.toObject(),
              });
            } catch (error) {
              console.log(
                `Forward history for internal memo with ID ${findInternalMemo?._id} is not created successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en: `An internal memo is CC'd to you by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}.`,
            Message_am: `የዉስጥ ማስታወሻ ደብዳቤ (ኦፊሻል የሆነ) ከ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ለእርስዎ CC ተደርጓል`,
          };

          await Notification.create({
            office_user: internalCC?.internal_office,
            notifcation_type: "InternalMemo",
            document_id: findInternalMemo?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(internalCC?.internal_office, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("internal_memo_forward_notification", {
              Message_en: `Internal memo is CC'd to you from ${
                findRequesterOfficeUser?.firstname +
                " " +
                findRequesterOfficeUser?.middlename +
                " " +
                findRequesterOfficeUser?.lastname +
                "-" +
                findRequesterOfficeUser?.position
              }.`,
              Message_am: `የዉስጥ ማስታወሻ ደብዳቤው ወደ እርስዎ ከ${
                findRequesterOfficeUser?.firstname +
                " " +
                findRequesterOfficeUser?.middlename +
                " " +
                findRequesterOfficeUser?.lastname +
                "-" +
                findRequesterOfficeUser?.position
              } CC ተደርጓል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `The internal memo is successfully verified by ${
            findRequesterOfficeUser?.firstname +
            " " +
            findRequesterOfficeUser?.middlename +
            " " +
            findRequesterOfficeUser?.lastname +
            "-" +
            findRequesterOfficeUser?.position
          }`,
          Message_am: `የዉስጥ ማስታወሻዉ በ${
            findRequesterOfficeUser?.firstname +
            " " +
            findRequesterOfficeUser?.middlename +
            " " +
            findRequesterOfficeUser?.lastname +
            "-" +
            findRequesterOfficeUser?.position
          } ስም ወጪ ሆኗል`,
        });
      } catch (error) {
        return res.status(StatusCodes.EXPECTATION_FAILED).json({
          Message_en:
            "The system is currently unable to generate the internal memo. Please try again.",
          Message_am:
            "ስርዓቱ በአሁኑ ጊዜ የዉስጥ ማስታወሻ ደብዳቤ ማመንጨት አልቻለም። እባክዎ ዳግም ይሞክሩ።",
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
  createInternalMemo,
  getInternalMemos,
  getInternalMemo,
  updInternalMemo,
  prvInternalMemo,
  vrfyInternalMemo,
};
