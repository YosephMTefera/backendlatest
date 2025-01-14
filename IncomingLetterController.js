const OfficeUser = require("../../model/OfficeUsers/OfficeUsers");
const LetterNumber = require("../../model/LetterNumbers/LetterNumber");
const ArchivalUser = require("../../model/ArchivalUsers/ArchivalUsers");
const IncomingLetter = require("../../model/IncomingLetters/IncomingLetter");
const IncomingLetterHistory = require("../../model/IncomingLetters/IncomingLetterHistory");

const { join } = require("path");
const cron = require("node-cron");
const mongoose = require("mongoose");
const formidable = require("formidable");
const { StatusCodes } = require("http-status-codes");
const { writeFile, readFile, unlink } = require("fs/promises");
const generateLetterNumber = require("../../utils/generateLetterNumber");
const getCurrentEthiopianDate = require("../../utils/getCurrentEthiopianDate");

const createIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_CRINCLETTER_API;
    const actualAPIKey = req?.headers?.get_crincletter_api;
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

      if (findRequesterArchivalUser?.status === "inactive") {
        return res.status(StatusCodes.UNAUTHORIZED).json({
          Message_en: "Not authorized to access data",
          Message_am: "ዳታውን ማግኘት አልተፈቀደሎትም",
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

        const attention_from = fields?.attention_from?.[0];
        const subject = fields?.subject?.[0];
        const sent_date = fields?.sent_date?.[0];
        const received_date = fields?.received_date?.[0];
        let no_attachment = fields?.no_attachment?.[0];
        const main_letter = files?.main_letter?.[0];
        const nimera = fields?.nimera?.[0];

        if (!attention_from) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Attention from is required",
            Message_am: "የጻፈውን ሰው ወይም መስሪያ ቤቱን ያስገቡ",
          });
        }

        if (!subject) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Subject is required",
            Message_am: "እባክዎን ጉዳዩን ያስገቡ",
          });
        }
        if (!sent_date) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please provide the date that the letter was sent",
            Message_am: "እባክዎ ደብዳቤው የተላከበትን ቀን ያስገቡ",
          });
        }
        if (!received_date) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please provide the date that the letter was received",
            Message_am: "እባክዎ ደብዳቤው የደረሰበትን ቀን ያስገቡ",
          });
        }
        if (!nimera) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please provide the nimera of the letter",
            Message_am: "እባክዎን የደብዳቤውን ንመራ ያስገቡ",
          });
        }
        if (!main_letter) {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en: "Please provide the letter itself",
            Message_am: "እባኮትን ደብዳቤውን ራሱ ያስገቡ",
          });
        }

        if (!no_attachment) {
          no_attachment = 0;
        }

        if (no_attachment) {
          if (isNaN(no_attachment)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "The number of attachment should only be a number",
              Message_am: "የአባሪዎች ብዛት ቁጥር ብቻ መሆን አለበት",
            });
          }

          if (no_attachment < 0) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "The number of attachments cannot be less than zero",
              Message_am: "የአባሪዎች ብዛት ከዜሮ ያነሰ መሆን አይችልም",
            });
          }
        }

        if (
          typeof main_letter === "object" &&
          (main_letter?.mimetype === "application/pdf" ||
            main_letter?.mimetype === "application/PDF")
        ) {
          if (main_letter?.size > 500 * 1024 * 1024) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Letter's file size is too large. Please insert a file less than 500MB",
              Message_am:
                "የደብዳቤው ፋይል መጠን በጣም ትልቅ ነው። እባክህ ከ 500ሜባ በታች የሆነ ፋይል አስገባ",
            });
          }
        } else {
          return res.status(StatusCodes.BAD_REQUEST).json({
            Message_en:
              "Invalid letter file format please try again. Only accepts '.pdf'",
            Message_am:
              "ልክ ያልሆነ የደብዳቤ ፋይል ቅርጸት እባክዎ እንደገና ይሞክሩ። «.pdf»ን ብቻ ይቀበላል",
          });
        }

        const letterNumber = await generateLetterNumber();

        const bytes = await readFile(main_letter?.filepath);
        const mainLetterBuffer = Buffer.from(bytes);
        const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);

        const path = join(
          "./",
          "Media",
          "IncomingLetterFiles",
          uniqueSuffix + "-" + main_letter?.originalFilename
        );

        const mainLetterAttachmentName =
          uniqueSuffix + "-" + main_letter?.originalFilename;

        await writeFile(path, mainLetterBuffer);

        const createIncomingLetter = await IncomingLetter.create({
          attention_from,
          subject,
          sent_date,
          received_date,
          no_attachment,
          main_letter: mainLetterAttachmentName,
          nimera,
          incoming_letter_number: letterNumber,
          createdBy: requesterId,
        });

        const updateHistory = [
          {
            updatedByArchivalUser: requesterId,
            action: "create",
          },
        ];

        try {
          await IncomingLetterHistory.create({
            incoming_letter_id: createIncomingLetter?._id,
            updateHistory,
            history: createIncomingLetter?.toObject(),
          });
        } catch (error) {
          console.log(
            `Incoming Letter history with this "${letterNumber}" letter number is not created`
          );
        }

        return res.status(StatusCodes.CREATED).json({
          Message_en: `Incoming letter number ${letterNumber} is created successfully`,
          Message_am: `የደብዳቤ ቁጥር ${letterNumber} ያለው ገቢ ደብዳቤ በተሳካ ሁኔታ ወደ ሲስተም ተመዝግቧል`,
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

cron.schedule("0 0 * * *", async () => {
  const [currentYear, currentMonth, currentDay] = getCurrentEthiopianDate();

  if (currentMonth === 11 && currentDay === 1) {
    let letterNumberRecord = await LetterNumber.findOne();

    if (!letterNumberRecord) {
      letterNumberRecord = await LetterNumber.create({});
    }

    letterNumberRecord.incoming_letter_number = 0;
    letterNumberRecord.outgoing_letter_number = 0;

    await letterNumberRecord.save();
  }
});

const getIncomingLetters = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INCLETTERS_API;
    const actualAPIKey = req?.headers?.get_incletters_api;
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

      let page = parseInt(req?.query?.page) || 1;
      let limit = parseInt(req?.query?.limit) || 10;
      let sortBy = parseInt(req?.query?.sort) || -1;
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

const getIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_INCLTRS_API;
    const actualAPIKey = req?.headers?.get_incltrs_api;
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

      const findIncomingLtr = await IncomingLetter.findOne({
        _id: id,
      }).populate({
        path: "createdBy",
        select: "_id firstname middlename lastname",
      });

      if (!findIncomingLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Incoming letter is not found",
          Message_am: "ገቢ ደብዳቤው አልተገኘም",
        });
      }

      return res.status(StatusCodes.OK).json(findIncomingLtr);
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

const updateIncomingLetter = async (req, res) => {
  try {
    const expectedURLKey = process?.env?.GET_UPDINCLTR_API;
    const actualAPIKey = req?.headers?.get_updincltr_api;

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

      if (findRequesterArchivalUser?.status === "inactive") {
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

      const findIncomingLtr = await IncomingLetter.findOne({ _id: id });

      if (!findIncomingLtr) {
        return res.status(StatusCodes.NOT_FOUND).json({
          Message_en: "Incoming letter is not found",
          Message_am: "ገቢ ደብዳቤው አልተገኘም",
        });
      }

      if (findIncomingLtr?.status === "forwarded") {
        return res.status(StatusCodes.FORBIDDEN).json({
          Message_en:
            "Incoming letter is already forwarded and cannot be edited",
          Message_am: "ገቢ ደብዳቤው ፎርዋርድ ተደርጓል እና ሊስተካከል አይችልም",
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

        const attention_from = fields?.attention_from?.[0];
        const subject = fields?.subject?.[0];
        const sent_date = fields?.sent_date?.[0];
        const received_date = fields?.received_date?.[0];
        let no_attachment = fields?.no_attachment?.[0];
        const main_letter = files?.main_letter?.[0];
        const nimera = fields?.nimera?.[0];

        const updatedFields = {};

        if (attention_from) {
          updatedFields.attention_from = attention_from;
        }
        if (subject) {
          updatedFields.subject = subject;
        }
        if (sent_date) {
          updatedFields.sent_date = sent_date;
        }
        if (received_date) {
          updatedFields.received_date = received_date;
        }
        if (no_attachment) {
          if (isNaN(no_attachment)) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "The number of attachment should only be a number",
              Message_am: "የአባሪዎች ብዛት ቁጥር ብቻ መሆን አለበት",
            });
          }

          if (no_attachment < 0) {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en: "The number of attachments cannot be less than zero",
              Message_am: "የአባሪዎች ብዛት ከዜሮ ያነሰ መሆን አይችልም",
            });
          }
          updatedFields.no_attachment = no_attachment;
        }
        if (nimera) {
          updatedFields.nimera = nimera;
        }
        if (main_letter) {
          if (
            typeof main_letter === "object" &&
            (main_letter?.mimetype === "application/pdf" ||
              main_letter?.mimetype === "application/PDF")
          ) {
            if (main_letter?.size > 500 * 1024 * 1024) {
              return res.status(StatusCodes.BAD_REQUEST).json({
                Message_en:
                  "Letter's file size is too large. Please insert a file less than 500MB",
                Message_am:
                  "የደብዳቤው ፋይል መጠን በጣም ትልቅ ነው። እባክህ ከ 500ሜባ በታች የሆነ ፋይል አስገባ",
              });
            }
          } else {
            return res.status(StatusCodes.BAD_REQUEST).json({
              Message_en:
                "Invalid letter file format please try again. Only accepts '.pdf'",
              Message_am:
                "ልክ ያልሆነ የደብዳቤ ፋይል ቅርጸት እባክዎ እንደገና ይሞክሩ። «.pdf»ን ብቻ ይቀበላል",
            });
          }

          const bytes = await readFile(main_letter?.filepath);
          const mainLetterBuffer = Buffer.from(bytes);
          const uniqueSuffix =
            Date.now() + "-" + Math.round(Math.random() * 1e9);

          const newPath = join(
            "./",
            "Media",
            "IncomingLetterFiles",
            uniqueSuffix + "-" + main_letter?.originalFilename
          );

          const mainLetterAttachmentName =
            uniqueSuffix + "-" + main_letter?.originalFilename;

          await writeFile(newPath, mainLetterBuffer);

          updatedFields.main_letter = main_letter
            ? mainLetterAttachmentName
            : findIncomingLtr?.main_letter;

          if (findIncomingLtr?.main_letter) {
            const path = join(
              "./",
              "Media",
              "IncomingLetterFiles",
              findIncomingLtr?.main_letter
            );

            try {
              await unlink(path);
            } catch (error) {
              console.log(
                `Incoming letter ${findIncomingLtr?.incoming_letter_number}'s attachment is replaced, but previous attachment is not found`
              );
            }
          }
        }

        const newUpdatedIncomingLtr = await IncomingLetter.findOneAndUpdate(
          { _id: id },
          updatedFields,
          { new: true }
        );

        if (!newUpdatedIncomingLtr) {
          return res.status(StatusCodes.NOT_FOUND).json({
            Message_en: "Incoming letter is not found",
            Message_am: "ገቢ ደብዳቤው አልተገኘም",
          });
        }

        try {
          await IncomingLetterHistory.findOneAndUpdate(
            { incoming_letter_id: id },
            {
              $push: {
                updateHistory: {
                  updatedByArchivalUser: requesterId,
                  action: "update",
                },
                history: newUpdatedIncomingLtr?.toObject(),
              },
            }
          );
        } catch (error) {
          console.log(
            `Incoming letter history with this "${findIncomingLtr?.incoming_letter_number}" letter number is not updated successfully`
          );
        }

        return res.status(StatusCodes.OK).json({
          Message_en: `Incoming letter number :- ${newUpdatedIncomingLtr?.incoming_letter_number} is updated successfully`,
          Message_am: `የገቢ ደብዳቤ ቁጥር፡- ${newUpdatedIncomingLtr?.incoming_letter_number} በተሳካ ሁኔታ ተስተካክሏል`,
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

module.exports = {
  createIncomingLetter,
  getIncomingLetters,
  getIncomingLetter,
  updateIncomingLetter,
};
