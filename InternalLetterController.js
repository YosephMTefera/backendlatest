const Division = require("../../model/Divisions/Divisions");
const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const TeamLeader = require("../../model/TeamLeaders/TeamLeaders");
const Directorate = require("../../model/Directorates/Directorates");
const Notification = require("../../model/Notifications/Notification");
const ArchivalUser = require("../../model/ArchivalUsers/ArchivalUsers");
const InternalLetter = require("../../model/InternalLetters/InternalLetter");
const InternalLetterHistory = require("../../model/InternalLetters/InternalLetterHistory");
const ForwardInternalLetter = require("../../model/ForwardInternalLetters/ForwardInternalLetter");
const ForwardInternalLetterHistory = require("../../model/ForwardInternalLetters/ForwardInternalLetterHistory");

const fs = require("fs");
const { join } = require("path");
const mongoose = require("mongoose");
const formidable = require("formidable");
var ethiopianDate = require("ethiopian-date");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile, unlink } = require("fs/promises");
const generateOutgoingLtrNo = require("../../utils/generateOutgoingLtrNo");
const {
  previewInternalLtrResponse,
} = require("../../middleware/previewInternalLtr");
const {
  previewInternalLtrOutput,
} = require("../../middleware/internalLtrOutput");
const { finalInternalLetter } = require("../../middleware/finalInputLtr");

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

const createInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CRINTLTR_API;
    const actualAPIKey = req?.headers?.get_crintltr_api;
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
        const output_by = fields?.output_by?.[0];
        const main_letter_attachment = files?.main_letter_attachment?.[0];

        if (!to_whom || to_whom?.length === 0) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please specify to whom the letter is written",
            Message_am: "ደብዳቤው ለማን እንደተጻፈ እባክዎ ይግለጹ",
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
            "InternalLetterAttachmentFiles",
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename
          );

          attachmentName =
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename;

          await writeFile(path, letterReplyAttachmentBuffer);
        }

        const createInternalLtr = await InternalLetter.create({
          to_whom: forwardToWhomArray?.map((item) => ({
            internal_office: item?.internal_office,
          })),
          to_whom_col: to_whom_col,
          subject: subject,
          body: body,
          internal_cc: forwardIntCC?.map((item) => ({
            internal_office: item?.internal_office,
          })),
          output_by: output_by,
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
          await InternalLetterHistory.create({
            internal_letter_id: createInternalLtr?._id,
            updateHistory,
            history: createInternalLtr?.toObject(),
          });
        } catch (error) {
          console.log(
            `Internal letter history with this ID ${createInternalLtr?._id} is not created`
          );
        }

        return res.status(StatusCodes.CREATED).json({
          Message_en: `The internal letter is created successfully.`,
          Message_am: `ውስጣዊ ደብዳቤዉ በተሳካ ሁኔታ ተፈጥሯል።`,
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

const getInternalLetters = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INTLTRS_API;
    const actualAPIKey = req?.headers?.get_intltrs_api;
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

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || -1;
      let internalLtrNum = req?.query?.internal_letter_number || "";
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

      const query = {};

      if (internalLtrNum) {
        query.internal_letter_number = {
          $regex: internalLtrNum,
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
      if (findRequesterArchivalUser) {
        query.$or = [{ status: "output" }, { status: "verified" }];
      }

      const totalInternalLtrs = await InternalLetter.countDocuments(query);

      const totalPages = Math.ceil(totalInternalLtrs / limit);

      if (page > totalPages) {
        page = 1;
      }

      const skip = (page - 1) * limit;

      const findInternalLtrs = await InternalLetter.find(query)
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
          path: "to_whom.internal_office",
          select: "_id firstname middlename lastname username position",
        })
        .populate({
          path: "internal_cc.internal_office",
          select: "_id firstname middlename lastname position",
        });

      if (!findInternalLtrs) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letters are not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤዎች አልተገኙም",
        });
      }

      return res.status(StatusCodes.OK).json({
        internalLetters: findInternalLtrs,
        totalInternalLtrs,
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

const getInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INTLTR_API;
    const actualAPIKey = req?.headers?.get_intltr_api;
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
      const findRequesterOfficeUser = await OfficeUser.findOne({
        _id: requesterId,
      });

      if (!findRequesterOfficeUser && !findRequesterArchivalUser) {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      if (findRequesterArchivalUser) {
        if (findRequesterArchivalUser?.status === "inactive") {
          return res.status(StatusCodes.UNAUTHORIZED).json({
            Message_en: "Not authorized to access data",
            Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
          });
        }
      }

      if (findRequesterOfficeUser) {
        if (findRequesterOfficeUser?.status === "inactive") {
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

      const findInternalLtr = await InternalLetter.findOne({ _id: id })
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
          path: "to_whom.internal_office",
          select: "_id firstname middlename lastname username position",
        })
        .populate({
          path: "internal_cc.internal_office",
          select: "_id firstname middlename lastname position",
        });

      if (!findInternalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letter is not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
        });
      }

      return res.status(StatusCodes.OK).json(findInternalLtr);
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

const updateInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INTLTRUPD_API;
    const actualAPIKey = req?.headers?.get_intltrupd_api;
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
        if (findRequesterOfficeUser?.status === "inactive") {
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

      const findInternalLtr = await InternalLetter.findOne({ _id: id });

      if (!findInternalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letter is not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalLtr?.status !== "pending") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated. So, this letter cannot be updated.",
          Message_am: "ደብዳቤው አስቀድሞ ተፈጥሯል። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      const findForwardInternalLtr = await ForwardInternalLetter.findOne({
        internal_letter_id: id,
      });

      const findForwardedPerson = findForwardInternalLtr?.path?.find(
        (item) =>
          item?.to?.toString() === requesterId?.toString() && item?.cc === "no"
      );

      if (
        findInternalLtr?.createdBy?.toString() !== requesterId?.toString() &&
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

        if (output_by) {
          if (
            findInternalLtr?.createdBy?.toString() === requesterId?.toString()
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

            updatedFields.output_by = output_by;
          }
        }

        if (detachFile) {
          if (detachFile === "detachAttachment") {
            if (findInternalLtr?.main_letter_attachment) {
              const detPath = join(
                "./",
                "Media",
                "InternalLetterAttachmentFiles",
                findInternalLtr?.main_letter_attachment
              );

              try {
                await unlink(detPath);
                findInternalLtr.main_letter_attachment = "";
                await findInternalLtr.save();
              } catch (error) {
                console.log(
                  `Internal letter with ID ${findInternalLtr?._id} attachment is not found`
                );
              }
            }
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
                  findInternalLtr?.internal_cc &&
                  findInternalLtr?.internal_cc?.length > 0
                ) {
                  for (const intUser of findInternalLtr?.internal_cc) {
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
                  findInternalLtr?.to_whom &&
                  findInternalLtr?.to_whom?.length > 0
                ) {
                  for (const toWhom of findInternalLtr?.to_whom) {
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
              findInternalLtr.to_whom,
              toWhomUpdates,
              updatedFields,
              "to_whom"
            );
          }

          if (fields?.internal_cc?.[0]) {
            const internalCcUpdates = JSON.parse(fields?.internal_cc?.[0]);
            await updateArrayField(
              findInternalLtr.internal_cc,
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
            "InternalLetterAttachmentFiles",
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename
          );

          const mainLetterAttachmentName =
            uniqueSuffix + "-" + main_letter_attachment?.originalFilename;

          await writeFile(newPath, mainLetterBuffer);

          updatedFields.main_letter_attachment = main_letter_attachment
            ? mainLetterAttachmentName
            : findInternalLtr?.main_letter_attachment;

          if (findInternalLtr?.main_letter_attachment) {
            const path = join(
              "./",
              "Media",
              "InternalLetterAttachmentFiles",
              findInternalLtr?.main_letter_attachment
            );

            try {
              await unlink(path);
            } catch (error) {
              console.log(
                `Internal letter with ID ${findInternalLtr?._id} attachment is replaced, but previous attachment is not found`
              );
            }
          }
        }

        const updatedByLists = [];

        if (
          findInternalLtr?.updated_by &&
          Array.isArray(findInternalLtr?.updated_by)
        ) {
          for (const existingUpdated of findInternalLtr?.updated_by) {
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

        const newUpdatedInternalLtr = await InternalLetter.findOneAndUpdate(
          { _id: id },
          updatedFields,
          { new: true }
        );

        if (!newUpdatedInternalLtr) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: "Internal letter is not found",
            Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
          });
        }

        try {
          await InternalLetterHistory.findOneAndUpdate(
            { internal_letter_id: id },
            {
              $push: {
                updateHistory: {
                  updatedByOfficeUser: requesterId,
                  action: "update",
                },
                history: newUpdatedInternalLtr?.toObject(),
              },
            }
          );
        } catch (error) {
          console.log(
            `Internal letter history with this ID ${findInternalLtr?._id} is not updated successfully`
          );
        }

        return res.status(StatusCodes.OK).json({
          Message_en: "Internal letter is updated successfully",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤዉ በተሳካ ሁኔታ ዘምኗል",
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

const previewInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_PRVINTLTR_API;
    const actualAPIKey = req?.headers?.get_prvintltr_api;
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
        if (findRequesterOfficeUser?.status === "inactive") {
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

      const findInternalLtr = await InternalLetter.findOne({ _id: id });

      if (!findInternalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letter is not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalLtr?.status !== "pending") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated. So, this letter cannot be previewed.",
          Message_am: "ደብዳቤው አስቀድሞ ተፈጥሯል። ስለዚህ ይህ ደብዳቤ በቅድመ-እይታ ሊታይ አይችልም።",
        });
      }

      if (findInternalLtr?.to_whom?.length === 0) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the list of recipients to this letter before previewing this letter",
          Message_am:
            "እባክዎ ይህንን ደብዳቤ አስቀድመው ከማየትዎ በፊት የዚህን ደብዳቤ የተቀባዮች ዝርዝር ይግለጹ",
        });
      }

      if (!findInternalLtr?.subject) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the subject of the letter before previewing this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ አስቀድመው ከማየትዎ በፊት የደብዳቤውን ርዕሰ ጉዳይ ይግለጹ",
        });
      }

      if (!findInternalLtr?.body) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the body of the letter before previewing this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ አስቀድመው ከማየትዎ በፊት የደብዳቤውን ሃተታ ያስገቡ",
        });
      }

      let lstOfToWhom = [];
      if (findInternalLtr?.to_whom?.length > 0) {
        for (const singlePerson of findInternalLtr?.to_whom) {
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
      if (findInternalLtr?.internal_cc?.length > 0) {
        for (const singlePerson of findInternalLtr?.internal_cc) {
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

      if (findInternalLtr?.main_letter_attachment) {
        attachmentPath = join(
          "./",
          "Media",
          "InternalLetterAttachmentFiles",
          findInternalLtr?.main_letter_attachment
        );
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
        "PreviewInternalLetters",
        uniqueSuffix + "-previewinternalltrs.pdf"
      );

      const text = [
        { toWhom: lstOfToWhom },
        { subject: findInternalLtr?.subject },
        { body: findInternalLtr?.body },
        { toExt: lstOfIntCC },
        { toWhomColNum: findInternalLtr?.to_whom_col },
        { internalCCNum: findInternalLtr?.internal_cc_col },
        { attachmentFile: attachmentPath },
      ];

      try {
        await previewInternalLtrResponse(inputPath, text, outputPath);

        const modifiedOutputPath = outputPath
          .replace(/\\/g, "/")
          .replace(/Media\//, "/");

        return res.status(StatusCodes.OK).json(modifiedOutputPath);
      } catch (error) {
        return res.status(StatusCodes.EXPECTATION_FAILED).json({
          Message_en:
            "The system is currently unable to generate the internal letter preview. Please try again.",
          Message_am:
            "ስርዓቱ በአሁኑ ጊዜ የዉስጥ ለዉስጥ ደብዳቤ ቅድመ እይታ መፍጠር አልቻለም። እባክዎ ዳግም ይሞክሩ።",
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

const outputInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_APPINTLTR_API;
    const actualAPIKey = req?.headers?.get_appintltr_api;
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

      if (findRequesterOfficeUser?.status !== "active") {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const id = req?.params?.id;

      if (!id || !mongoose.isValidObjectId(id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findInternalLtr = await InternalLetter.findOne({ _id: id });

      if (!findInternalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letter is not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalLtr?.status !== "pending") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated. So, this letter cannot be updated.",
          Message_am: "ደብዳቤው አስቀድሞ ተፈጥሯል። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      if (findInternalLtr?.output_by?.toString() !== requesterId?.toString()) {
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

      if (findInternalLtr?.to_whom?.length === 0) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the list of recipients to this letter before approving this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ ወጪ ከማድረግዎ በፊት የዚህን ደብዳቤ የተቀባዮች ዝርዝር ይግለጹ",
        });
      }

      if (!findInternalLtr?.subject) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the subject of the letter before approving this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ ወጪ ከማድረግዎ በፊት የደብዳቤውን ርዕሰ ጉዳይ ይግለጹ",
        });
      }

      if (!findInternalLtr?.body) {
        return res.status(StatusCodes.BAD_REQUEST).json({
          Message_en:
            "Please specify the body of the letter before approving this letter",
          Message_am: "እባክዎ ይህንን ደብዳቤ ወጪ ከማድረግዎ በፊት የደብዳቤውን ሃተታ ያስገቡ",
        });
      }

      for (const lstOfAcceptors of findInternalLtr?.to_whom) {
        const findAcceptorUser = await OfficeUser.findOne({
          _id: lstOfAcceptors?.internal_office,
        });

        if (!findAcceptorUser) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: "The receiver of this letter is not found",
            Message_am: "የዚህ ደብዳቤ ተቀባይ አልተገኘም",
          });
        }

        if (findAcceptorUser?.status === "inactive") {
          return res.status(StatusCodes.FORBIDDEN).json({
            Message_en: `The user to receive the letter is currently inactive. (${
              findAcceptorUser?.firstname +
              " " +
              findAcceptorUser?.middlename +
              " " +
              findAcceptorUser?.lastname
            })`,
            Message_am: `የዚህን ደብዳቤ ተቀባይ ተጠቃሚ በአሁኑ ጊዜ አክቲቭ አይደለም። (${
              findAcceptorUser?.firstname +
              " " +
              findAcceptorUser?.middlename +
              " " +
              findAcceptorUser?.lastname
            })`,
          });
        }

        if (findAcceptorUser?.level === "MainExecutive") {
          const findRequesterDivision = await Division.findOne({
            _id: findRequesterOfficeUser?.division,
          });

          if (
            findRequesterDivision?.special !== "yes" &&
            findRequesterOfficeUser?.level !== "DivisionManagers"
          ) {
            return res.status(StatusCodes.UNAUTHORIZED).json({
              Message_en: `You are not authorized to write an internal letter to the main director. (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
              Message_am: `ለዋናው ዳይፌክተር የውስጥ ደብዳቤ መጻፍ አልተፈቀደልዎትም። (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
            });
          }
        }

        if (findAcceptorUser?.level === "DivisionManagers") {
          if (findRequesterOfficeUser?.level === "Directors") {
            if (
              findRequesterOfficeUser?.division?.toString() !==
              findAcceptorUser?.division?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `A director cannot write a letter to a division manager other than his own division manager. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `አንድ ዳይሬክተር ከራሱ ዘርፍ ኃላፊ ውጪ ለሌላ ዘርፍ ኃላፊ ደብዳቤ መጻፍ አይችልም። (${
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
            findRequesterOfficeUser?.level === "TeamLeaders" ||
            findRequesterOfficeUser?.level === "Professionals"
          ) {
            return res.status(StatusCodes.UNAUTHORIZED).json({
              Message_en: `A team leader or professional is not authorized to write an internal letter to a division manager. (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
              Message_am: `የቡድን መሪ ወይም ባለሙያ ለዘርፍ ኃላፊ የውስጥ ደብዳቤ ለመጻፍ አልተፈቀደለትም። (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
            });
          }
        }

        if (findAcceptorUser?.level === "Directors") {
          if (findRequesterOfficeUser?.level === "DivisionManagers") {
            if (
              findRequesterOfficeUser?.division?.toString() !==
              findAcceptorUser?.division?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `A division manager cannot write an internal letter to directors in another division. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የዘርፍ ኃላፊ በሌላ ዘርፍ ውስጥ ላሉ ዳይሬክተሮች የውስጥ ደብዳቤ መጻፍ አይችልም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "TeamLeaders") {
            const findTeamLeadersDirectorate = await Directorate.findOne({
              "members.users": findRequesterOfficeUser?._id,
            });

            if (!findTeamLeadersDirectorate) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `You are not part of any directorate, thus you cannot write a letter to Director ${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                }.`,
                Message_am: `የየትኛዉም ዳይሬክቶሬት አባል ስላልሆኑ ፤ ለዳይሬክተር ${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                } ደብዳቤ መጻፍ አይችሉም።`,
              });
            }

            if (
              findTeamLeadersDirectorate?.manager?.toString() !==
              findAcceptorUser?._id?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `You are attempting to write a letter to a directorate you are not part of. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `እርስዎ አባል ላልሆኑበት ዳይሬክቶሬት ደብዳቤ ለመጻፍ እየሞከሩ ነው። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "Professionals") {
            return res.status(StatusCodes.UNAUTHORIZED).json({
              Message_en: `A professional is not authorized to write an internal letter to a director. (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
              Message_am: `ባለሙያ ለዳይሬክተር የውስጥ ደብዳቤ መጻፍ አልተፈቀደለትም። (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
            });
          }
        }
        if (findAcceptorUser?.level === "TeamLeaders") {
          if (findRequesterOfficeUser?.level === "DivisionManagers") {
            if (
              findRequesterOfficeUser?.division?.toString() !==
              findAcceptorUser?.division?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `A division manager cannot write an internal letter to teams in another division. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የዘርፍ ኃላፊ በሌላ ዘርፍ ውስጥ ላሉ ቡድኖች የውስጥ ደብዳቤ መጻፍ አይችልም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "Directors") {
            const findTeamLeadersDirectorate = await Directorate.findOne({
              "members.users": findAcceptorUser?._id,
            });

            if (!findTeamLeadersDirectorate) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The team is not part of your directorate, thus you cannot write a letter to Team leader ${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                }.`,
                Message_am: `ቡድኑ የእርስዎ ዳይሬክቶሬት አባል ስላልሆነ ፤ ለቡድን መሪ ${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                } ደብዳቤ መጻፍ አይችሉም።`,
              });
            }

            if (
              findTeamLeadersDirectorate?.manager?.toString() !==
              findRequesterOfficeUser?._id?.toString()
            ) {
              return res.status(StatusCodes.NOT_FOUND).json({
                Message_en: `The team is not part of your directorate, thus you cannot write a letter to Team leader ${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                }.`,
                Message_am: `ቡድኑ የእርስዎ ዳይሬክቶሬት አባል ስላልሆነ ፤ ለቡድን መሪ ${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                } ደብዳቤ መጻፍ አይችሉም።`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "TeamLeaders") {
            if (
              findRequesterOfficeUser?.division?.toString() !==
              findAcceptorUser?.division?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `You cannot write a letter to a team in another division. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `በሌላ ዘርፍ ውስጥ ላለ ቡድን ደብዳቤ መጻፍ አይችሉም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "Professionals") {
            const findProfessionalsInTeam = await TeamLeader.findOne({
              "members.users": findRequesterOfficeUser?._id,
            });

            if (!findProfessionalsInTeam) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `Professionals cannot write letters to teams of which they are not members. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `ባለሙያዎች አባል ላልሆኑበት ቡድን ደብዳቤ ሊጽፉ አይችሉም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }

            if (
              findProfessionalsInTeam?.manager?.toString() !==
              findAcceptorUser?._id?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `Professionals cannot write letters to teams of which they are not members. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `ባለሙያዎች አባል ላልሆኑበት ቡድን ደብዳቤ ሊጽፉ አይችሉም። (${
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

        if (findAcceptorUser?.level === "Professionals") {
          if (findRequesterOfficeUser?.level === "DivisionManagers") {
            if (
              findRequesterOfficeUser?.division?.toString() !==
              findAcceptorUser?.division?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `A division manager cannot write an internal letter to professionals in another division. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የዘርፍ ኃላፊ በሌላ ዘርፍ ውስጥ ላሉ ባለሙያዎች የውስጥ ደብዳቤ መጻፍ አይችልም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "Directors") {
            const findProfessionalsInDirectorates = await Directorate.findOne({
              "members.users": findAcceptorUser?._id,
            });

            if (!findProfessionalsInDirectorates) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `You cannot write an internal letter to a professional that is not part of your directorate. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የዳይሬክቶሬትዎ አካል ላልሆነ ባለሙያ የውስጥ ደብዳቤ መጻፍ አይችሉም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }

            if (
              findProfessionalsInDirectorates?.manager?.toString() !==
              findRequesterOfficeUser?._id?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `You cannot write an internal letter to a professional that is not part of your directorate. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የዳይሬክቶሬትዎ አካል ላልሆነ ባለሙያ የውስጥ ደብዳቤ መጻፍ አይችሉም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "TeamLeaders") {
            const findProfessionalsInTeam = await TeamLeader.findOne({
              "members.users": findAcceptorUser?._id,
            });

            if (!findProfessionalsInTeam) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `You cannot write an internal letter to a professional that is not part of your team. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የቡድንዎ አካል ላልሆነ ባለሙያ የውስጥ ደብዳቤ መጻፍ አይችሉም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }

            if (
              findProfessionalsInTeam?.manager?.toString() !==
              findRequesterOfficeUser?._id?.toString()
            ) {
              return res.status(StatusCodes.UNAUTHORIZED).json({
                Message_en: `You cannot write an internal letter to a professional that is not part of your team. (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
                Message_am: `የቡድንዎ አካል ላልሆነ ባለሙያ የውስጥ ደብዳቤ መጻፍ አይችሉም። (${
                  findAcceptorUser?.firstname +
                  " " +
                  findAcceptorUser?.middlename +
                  " " +
                  findAcceptorUser?.lastname
                })`,
              });
            }
          }

          if (findRequesterOfficeUser?.level === "Professionals") {
            return res.status(StatusCodes.UNAUTHORIZED).json({
              Message_en: `A professional cannot write an internal letter to another professional. (${
                findAcceptorUser?.firstname +
                " " +
                findAcceptorUser?.middlename +
                " " +
                findAcceptorUser?.lastname
              })`,
              Message_am: `አንድ ባለሙያ ለሌላ ባለሙያ ውስጣዊ ደብዳቤ መጻፍ አይችልም። (${
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

      let lstOfToWhom = [];
      if (findInternalLtr?.to_whom?.length > 0) {
        for (const singlePerson of findInternalLtr?.to_whom) {
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
      if (findInternalLtr?.internal_cc?.length > 0) {
        for (const singlePerson of findInternalLtr?.internal_cc) {
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

      // Check if the Main Director is CC'd with in the Division Managers exchange
      if (findRequesterOfficeUser?.level === "DivisionManagers") {
        let containsAnotherDivisionManager = false;
        for (const person of findInternalLtr?.to_whom) {
          const user = await OfficeUser.findOne({
            _id: person?.internal_office,
          });
          if (user?.level === "DivisionManagers") {
            containsAnotherDivisionManager = true;
            break;
          }
        }

        if (containsAnotherDivisionManager) {
          let hasMainExec = false;
          for (const person of lstOfIntCC) {
            const user = await OfficeUser.findOne({
              position: person?.internal_office,
            });
            if (user?.level === "MainExecutive") {
              hasMainExec = true;
              break;
            }
          }

          if (!hasMainExec) {
            const mainExec = await OfficeUser.findOne({
              level: "MainExecutive",
              status: "active",
            });
            if (mainExec) {
              findInternalLtr.internal_cc.push({
                internal_office: mainExec?._id,
              });
              await findInternalLtr.save();
              lstOfIntCC.push({ internal_office: mainExec?.position });
            }
          }
        }
      }

      // Check for DivisionManager if requester is Director and to_whom has another Director from another division
      if (findRequesterOfficeUser?.level === "Directors") {
        let containsAnotherDirectorFromAnotherDivision = false;
        for (const person of findInternalLtr?.to_whom) {
          const user = await OfficeUser.findOne({
            _id: person?.internal_office,
          });
          if (
            user?.level === "Directors" &&
            user?.division !== findRequesterOfficeUser?.division
          ) {
            containsAnotherDirectorFromAnotherDivision = true;
            break;
          }
        }

        if (containsAnotherDirectorFromAnotherDivision) {
          let hasDivisionManager = false;
          for (const person of lstOfIntCC) {
            const user = await OfficeUser.findOne({
              position: person?.internal_office,
            });
            if (
              user?.level === "DivisionManagers" &&
              user?.division === findRequesterOfficeUser?.division
            ) {
              hasDivisionManager = true;
              break;
            }
          }

          if (!hasDivisionManager) {
            const divisionManager = await OfficeUser.findOne({
              level: "DivisionManagers",
              division: findRequesterOfficeUser?.division,
              status: "active",
            });
            if (divisionManager) {
              findInternalLtr.internal_cc.push({
                internal_office: divisionManager?._id,
              });
              await findInternalLtr.save();
              lstOfIntCC.push({ internal_office: divisionManager?.position });
            }
          }
        }
      }

      // Sort lstOfIntCC based on organizational structure
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
      if (findInternalLtr?.main_letter_attachment) {
        attachmentPath = join(
          "./",
          "Media",
          "InternalLetterAttachmentFiles",
          findInternalLtr?.main_letter_attachment
        );
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
        "InternalLetterFiles",
        uniqueSuffix + "-internalltrs.pdf"
      );

      const text = [
        { toWhom: lstOfToWhom },
        { subject: findInternalLtr?.subject },
        { body: findInternalLtr?.body },
        { verSign: checkSignature },
        { verTiter: checkTiter },
        { toWhomColNum: findInternalLtr?.to_whom_col },
        { internalCCNum: findInternalLtr?.internal_cc_col },
        { toExt: lstOfIntCCWithLevels },
        { attachmentFile: attachmentPath },
      ];

      try {
        await previewInternalLtrOutput(inputPath, text, outputPath);

        const modifiedOutputPath = outputPath
          .replace(/\\/g, "/")
          .replace(/Media\//, "/");

        const updatedFields = {};
        updatedFields.output_date = new Date();
        updatedFields.status = "output";
        updatedFields.main_letter = modifiedOutputPath;

        const internalLtr = await InternalLetter.findOneAndUpdate(
          { _id: id },
          updatedFields,
          { new: true }
        );

        if (!internalLtr) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: `Internal letter is not found`,
            Message_am: `የዉስጥ ደብዳቤዉ አልተገኘም`,
          });
        }

        try {
          await InternalLetterHistory.findOneAndUpdate(
            { internal_letter_id: id },
            {
              $push: {
                updateHistory: {
                  updatedByOfficeUser: requesterId,
                  action: "update",
                },
                history: internalLtr?.toObject(),
              },
            }
          );

          const io = global?.io;
          const onlineUserList = global?.onlineUserList;

          const notificationMessage = {
            Message_en: `The internal letter with subject ${findInternalLtr?.subject} is submitted to the archivals, to be verified by the archivals.`,
            Message_am: `የደብዳቤ ርዕስ ${findInternalLtr?.subject} ያለዉ የወስጥ ደብዳቤ ወደ መዝገብ ቤት ወጪ እንዲሆን ተልኳል።`,
          };

          const findArchivals = await ArchivalUser.find({
            status: "active",
          });

          if (findArchivals?.length > 0) {
            for (const archs of findArchivals) {
              await Notification.create({
                archival_user: archs?._id,
                notifcation_type: "InternalLetter",
                document_id: findInternalLtr?._id,
                message_en: notificationMessage?.Message_en,
                message_am: notificationMessage?.Message_am,
              });

              const user = getUser(archs?._id, onlineUserList);
              if (user) {
                io.to(user?.socketID).emit("output_internal_ltr_notification", {
                  Message_en: `${notificationMessage?.Message_en}`,
                  Message_am: `${notificationMessage?.Message_am}`,
                });
              }
            }
          }
        } catch (error) {
          console.log(
            `Internal letter history with ID ${findInternalLtr?._id} is not updated successfully.`
          );
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `The internal letter is approved by ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname}`,
          Message_am: `ይህ የዉስጥ ደብዳቤ በ ${findRequesterOfficeUser?.firstname} ${findRequesterOfficeUser?.middlename} ${findRequesterOfficeUser?.lastname} ወጪ እንዲሆን ጸድቋል።`,
        });
      } catch (error) {
        return res.status(StatusCodes.EXPECTATION_FAILED).json({
          Message_en:
            "The system is currently unable to generate the internal letter. Please try again.",
          Message_am:
            "ስርዓቱ በአሁኑ ጊዜ የዉስጥ ለዉስጡን ደብዳቤ ማመንጨት አልቻለም። እባክዎ ዳግም ይሞክሩ።",
        });
      }
    } else {
      return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json();
    }
  } catch (error) {
    return res.status(StatusCodes.INTERNAL_SERVER_ERROR).json({
      Message_en: "Something went wrong please try again" + error?.message,
      Message_am: "ችግር ተፈጥሯል እባክዎ እንደገና ይሞክሩ",
    });
  }
};

const reverseInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_RVRSEINTLTRUPD_API;
    const actualAPIKey = req?.headers?.get_rvrseintltrupd_api;
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

      if (findRequesterOfficeUser?.status !== "active") {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
        });
      }

      const id = req?.params?.id;

      if (!id || !mongoose.isValidObjectId(id)) {
        return res.status(StatusCodes.NOT_ACCEPTABLE).json({
          Message_en: "Invalid request",
          Message_am: "ልክ ያልሆነ ጥያቄ",
        });
      }

      const findInternalLtr = await InternalLetter.findOne({ _id: id });

      if (!findInternalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letter is not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalLtr?.output_by?.toString() !== requesterId?.toString()) {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en: `The letter was not signed by you. Therefore, you cannot reverse the process.`,
          Message_am: `ደብዳቤዉ በእርሶ ስም ወጪ ስላልተደረገ ፤ የደብዳቤዉን ሂደት ወደ ኋላ መመለስ አይችሉም።`,
        });
      }

      if (findInternalLtr?.status !== "output") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated or still being processed. So, this letter cannot be updated.",
          Message_am:
            "ደብዳቤው አስቀድሞ ተፈጥሯል ወይም ደግሞ ገና እየተዘጋጀ ነዉ። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      const updatedFields = {};

      updatedFields.output_date = null;
      updatedFields.status = "pending";
      updatedFields.main_letter = "";

      const internalLtr = await InternalLetter.findOneAndUpdate(
        { _id: id },
        updatedFields,
        { new: true }
      );

      if (!internalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: `Internal letter is not found`,
          Message_am: `የዉስጥ ደብዳቤዉ አልተገኘም`,
        });
      }

      try {
        await InternalLetterHistory.findOneAndUpdate(
          { internal_letter_id: id },
          {
            $push: {
              updateHistory: {
                updatedByOfficeUser: requesterId,
                action: "update",
              },
              history: internalLtr?.toObject(),
            },
          }
        );

        if (findInternalLtr?.main_letter) {
          const oldPath = join("./", "Media", findInternalLtr?.main_letter);

          try {
            await unlink(oldPath);
          } catch (error) {
            console.log(
              `Internal letter with ID ${findInternalLtr?._id} attachment is replaced, but previous attachment is not found`
            );
          }
        }
      } catch (error) {
        console.log(
          `Internal letter history with ID ${findInternalLtr?._id} is not updated successfully.`
        );
      }

      return res.status(StatusCodes.OK).json({
        Message_en: "Internal letter is updated successfully",
        Message_am: "የዉስጥ ደብዳቤዉ በተሳካ ሁኔታ ዘምኗል",
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

const verifyInternalLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_VRFYINTLTR_API;
    const actualAPIKey = req?.headers?.get_vrfyintltr_api;
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

      const findInternalLtr = await InternalLetter.findOne({ _id: id });

      if (!findInternalLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Internal letter is not found",
          Message_am: "የዉስጥ ለዉስጥ ደብዳቤው አልተገኘም",
        });
      }

      if (findInternalLtr?.status !== "output") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "The letter is already generated or still being processed. So, this letter cannot be updated.",
          Message_am:
            "ደብዳቤው አስቀድሞ ተፈጥሯል ወይም ደግሞ ገና እየተዘጋጀ ነዉ። ስለዚህ ይህ ደብዳቤ ሊዘመን አይችልም።",
        });
      }

      const io = global?.io;
      const onlineUserList = global?.onlineUserList;
      const verified_date = new Date();
      let verDate = caseSubDate(verified_date);
      const findMainLtrAttachment = findInternalLtr?.main_letter;
      const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);

      const waterMarkPath = join(
        "./",
        "Media",
        "GenerateFiles",
        "WaterMark.png"
      );

      const outputPath = join(
        "./",
        "Media",
        "InternalLetterFiles",
        uniqueSuffix + "-finalinternalltrs.pdf"
      );

      const checkMainLetterExist = join("./", "Media", findMainLtrAttachment);

      if (!fs.existsSync(checkMainLetterExist)) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en:
            "The approved letter file is not found. This shows that it is manually manipulated. For this reason the letter cannot be verified.",
          Message_am:
            "የጸደቀው ደብዳቤ ፋይል አልተገኘም። ይህ የሚያሳየው ያለአግባብ እንደተነካ ነው። በዚህ ምክንያት ደብዳቤው ሲስተሙ ደብዳቤዉን ማተም አይችልም።",
        });
      }

      const letterNumber = await generateOutgoingLtrNo();

      const text = [
        { verDate: verDate },
        { waterMarkPath: waterMarkPath },
        { internalLtrNum: letterNumber },
        { finalInternalLtr: findMainLtrAttachment },
      ];

      try {
        await finalInternalLetter(text, outputPath);

        const modifiedOutputPath = outputPath
          .replace(/\\/g, "/")
          .replace(/Media\//, "/");

        const updatedFields = {};

        updatedFields.internal_letter_number = letterNumber;
        updatedFields.main_letter = modifiedOutputPath;
        updatedFields.verified_by = requesterId;
        updatedFields.verified_date = new Date();
        updatedFields.status = "verified";

        const verfyInternalLtr = await InternalLetter.findOneAndUpdate(
          { _id: id },
          updatedFields,
          { new: true }
        );

        if (!verfyInternalLtr) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: `Internal letter is not found`,
            Message_am: `የዉስጥ ደብዳቤዉ አልተገኘም`,
          });
        }

        try {
          await InternalLetterHistory.findOneAndUpdate(
            { internal_letter_id: id },
            {
              $push: {
                updateHistory: {
                  updatedByArchivalUser: requesterId,
                  action: "update",
                },
                history: verfyInternalLtr?.toObject(),
              },
            }
          );
        } catch (error) {
          console.log(
            `Internal letter with ID ${findInternalLtr?._id} is not updated successfully`
          );
        }

        for (const toWhom of findInternalLtr?.to_whom) {
          const findForwardInternalLtr = await ForwardInternalLetter.findOne({
            internal_letter_id: id,
          });

          if (findForwardInternalLtr) {
            const forwardInternalLtrToOfficer =
              await ForwardInternalLetter.findOneAndUpdate(
                { internal_letter_id: id },
                {
                  $push: {
                    path: {
                      forwarded_date: new Date(),
                      from_achival_user: requesterId,
                      cc: "no",
                      to: toWhom?.internal_office,
                    },
                  },
                },
                { new: true }
              );

            try {
              await ForwardInternalLetterHistory.findOneAndUpdate(
                {
                  forward_internal_letter_id: forwardInternalLtrToOfficer?._id,
                },
                {
                  $push: {
                    updateHistory: {
                      updatedByArchivalUser: requesterId,
                      action: "update",
                    },
                    history: forwardInternalLtrToOfficer?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Forward history for internal letter with ID ${findInternalLtr?._id} is not updated successfully`
              );
            }
          }

          if (!findForwardInternalLtr) {
            const path = [
              {
                forwarded_date: new Date(),
                from_achival_user: requesterId,
                cc: "no",
                to: toWhom?.internal_office,
              },
            ];

            const newInternalLetterForward = await ForwardInternalLetter.create(
              {
                internal_letter_id: id,
                path: path,
              }
            );

            const updateHistory = [
              {
                updatedByArchivalUser: requesterId,
                action: "create",
              },
            ];

            try {
              await ForwardInternalLetterHistory.create({
                forward_internal_letter_id: newInternalLetterForward?._id,
                updateHistory: updateHistory,
                history: newInternalLetterForward?.toObject(),
              });
            } catch (error) {
              console.log(
                `Forward history for internal letter with ID ${findInternalLtr?._id} is not created successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en: `An official internal letter is forwarded to you by ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (Archival).`,
            Message_am: `የዉስጥ ደብዳቤ (ኦፊሻል የሆነ) ከ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (መዝገብ ቤት) ለእርስዎ ተልኳል።`,
          };

          await Notification.create({
            office_user: toWhom?.internal_office,
            notifcation_type: "InternalLetter",
            document_id: findInternalLtr?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(toWhom?.internal_office, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("internal_letter_forward_notification", {
              Message_en: `Internal letter is forwarded to you from archivals.`,
              Message_am: `የዉስጥ ደብዳቤው ወደ እርስዎ ከመዝገብ ቤት ተልኳል።`,
            });
          }
        }

        for (const internalCC of findInternalLtr?.internal_cc) {
          const findForwardInternalLtr = await ForwardInternalLetter.findOne({
            internal_letter_id: id,
          });

          if (findForwardInternalLtr) {
            const forwardInternalLtrToOfficer =
              await ForwardInternalLetter.findOneAndUpdate(
                { internal_letter_id: id },
                {
                  $push: {
                    path: {
                      forwarded_date: new Date(),
                      from_achival_user: requesterId,
                      cc: "yes",
                      to: internalCC?.internal_office,
                    },
                  },
                },
                { new: true }
              );

            try {
              await ForwardInternalLetterHistory.findOneAndUpdate(
                {
                  forward_internal_letter_id: forwardInternalLtrToOfficer?._id,
                },
                {
                  $push: {
                    updateHistory: {
                      updatedByArchivalUser: requesterId,
                      action: "update",
                    },
                    history: forwardInternalLtrToOfficer?.toObject(),
                  },
                }
              );
            } catch (error) {
              console.log(
                `Forward history for internal letter with ID ${findInternalLtr?._id} is not updated successfully`
              );
            }
          }

          if (!findForwardInternalLtr) {
            const path = [
              {
                forwarded_date: new Date(),
                from_achival_user: requesterId,
                cc: "yes",
                to: internalCC?.internal_office,
              },
            ];

            const newInternalLetterForward = await ForwardInternalLetter.create(
              {
                internal_letter_id: id,
                path: path,
              }
            );

            const updateHistory = [
              {
                updatedByArchivalUser: requesterId,
                action: "create",
              },
            ];

            try {
              await ForwardInternalLetterHistory.create({
                forward_internal_letter_id: newInternalLetterForward?._id,
                updateHistory: updateHistory,
                history: newInternalLetterForward?.toObject(),
              });
            } catch (error) {
              console.log(
                `Forward history for internal letter with ID: ${findInternalLtr?._id}  is not created successfully`
              );
            }
          }

          const notificationMessage = {
            Message_en: `An official internal letter is CC'd to you by ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (Archival).`,
            Message_am: `የዉስጥ ደብዳቤ (ኦፊሻል የሆነ) ከ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} - (መዝገብ ቤት) ለእርስዎ CC ተደርጓል።`,
          };

          await Notification.create({
            office_user: internalCC?.internal_office,
            notifcation_type: "InternalLetter",
            document_id: findInternalLtr?._id,
            message_en: notificationMessage?.Message_en,
            message_am: notificationMessage?.Message_am,
          });

          const user = getUser(internalCC?.internal_office, onlineUserList);

          if (user) {
            io.to(user?.socketID).emit("internal_letter_forward_notification", {
              Message_en: `Internal letter is CC'd to you from archivals.`,
              Message_am: `የዉስጥ ደብዳቤው ወደ እርስዎ ከመዝገብ ቤት CC ተደርጓል።`,
            });
          }
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `The internal letter (${verfyInternalLtr?.internal_letter_number}) is successfully verified by ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname}.`,
          Message_am: `የዉስጥ ደብዳቤው (${verfyInternalLtr?.internal_letter_number})  በተሳካ ሁኔታ በ ${findRequesterArchivalUser?.firstname} ${findRequesterArchivalUser?.middlename} ${findRequesterArchivalUser?.lastname} ወጪ ተደርጓል።`,
        });
      } catch (error) {
        return res.status(StatusCodes.EXPECTATION_FAILED).json({
          Message_en:
            "The system is unable to generate the file for the internal letter. Please try again.",
          Message_am: "ሲስተሙ የወስጥ ለዉስጡን ደብዳቤ ማመንጨት አልቻለም። እባክዎ እንደገና ይሞክሩ።",
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
  createInternalLetter,
  getInternalLetters,
  getInternalLetter,
  previewInternalLetter,
  updateInternalLetter,
  reverseInternalLetter,
  outputInternalLetter,
  verifyInternalLetter,
};
