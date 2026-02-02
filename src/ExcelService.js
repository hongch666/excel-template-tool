const ExcelJS = require("exceljs");
const axios = require("axios");
const path = require("path");
const fs = require("fs");

class ExcelService {
  /**
   * 根据模板和数据生成 Excel 文件 Buffer
   * @param {Record<string, any>} data 需要填充到模板的数据对象
   * @param {string} templatePath 模板文件的绝对路径或相对路径
   * @returns {Promise<Buffer>} Excel 文件 Buffer
   */
  async exportToExcel(data, templatePath) {
    // 读取模板文件
    const absTemplatePath = path.isAbsolute(templatePath)
      ? templatePath
      : path.resolve(process.cwd(), templatePath);

    if (!fs.existsSync(absTemplatePath)) {
      throw new Error(`Template file not found: ${absTemplatePath}`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(absTemplatePath);
    const worksheet = workbook.getWorksheet(1);

    if (!worksheet) {
      throw new Error("Worksheet not found in template");
    }

    // 1. 填充基本信息
    await this.fillBasicInfo(worksheet, data);

    // 2. 填充数组数据
    await this.fillArrayData(worksheet, data);

    // 生成 Buffer
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  /**
   * 填充基本信息
   * @private
   */
  async fillBasicInfo(worksheet, data) {
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      for (let colNumber = 1; colNumber <= row.cellCount; colNumber++) {
        const cell = row.getCell(colNumber);

        // 修改判断条件，支持字符串和富文本对象
        if (
          cell.value &&
          (typeof cell.value === "string" || this.isRichText(cell.value))
        ) {
          let originalValue;

          // 获取文本内容
          if (typeof cell.value === "string") {
            originalValue = cell.value;
          } else {
            // 从富文本对象中提取纯文本
            originalValue = this.extractTextFromRichText(cell.value);
          }

          // 检查是否包含占位符
          const hasPlaceholder = /\$\{(\w+)\}/.test(originalValue);

          if (hasPlaceholder) {
            // 处理富文本单元格，保持格式
            await this.handleRichTextCell(
              cell,
              data,
              originalValue,
              worksheet,
              rowNumber,
              colNumber
            );
          } else {
            // 普通占位符替换
            if (typeof cell.value === "string") {
              let newValue = cell.value;
              Object.keys(data).forEach((key) => {
                if (
                  typeof data[key] !== "object" &&
                  data[key] !== null &&
                  data[key] !== undefined
                ) {
                  const regex = new RegExp(`^\\$\\{${key}\\}$`, "g");
                  newValue = newValue.replace(regex, String(data[key]));
                }
              });

              if (newValue !== originalValue) {
                cell.value = newValue;
              }
            }
          }
        }
      }
    }
  }

  /**
   * 填充数组数据
   * @private
   */
  async fillArrayData(worksheet, data) {
    const arrayTemplates = this.findArrayTemplates(worksheet);

    // 从后往前处理，避免行号变化影响
    for (const template of arrayTemplates.reverse()) {
      await this.processArrayTemplate(worksheet, template, data);
    }
  }

  /**
   * 处理富文本单元格，保持格式
   * @private
   */
  async handleRichTextCell(
    cell,
    data,
    originalText,
    worksheet,
    rowNumber,
    colNumber
  ) {
    // 分析占位符位置
    const placeholderRegex = /\$\{(\w+)\}/g;
    let match;
    const replacements = [];

    while ((match = placeholderRegex.exec(originalText)) !== null) {
      const key = match[1];
      if (
        data.hasOwnProperty(key) &&
        typeof data[key] !== "object" &&
        data[key] !== null &&
        data[key] !== undefined
      ) {
        const value = String(data[key]);

        // 检查是否为图片链接
        if (this.isImageUrl(value)) {
          await this.insertImageToCell(worksheet, rowNumber, colNumber, value);
          return; // 如果是图片，直接返回，不进行文本替换
        }

        // 无论是否为图片链接，都添加到替换列表中
        replacements.push({
          start: match.index,
          end: match.index + match[0].length,
          key: key,
          value: value,
        });
      }
    }

    if (replacements.length === 0) return;

    const cellValue = cell.value;

    // 如果是纯字符串，转换为富文本处理
    if (typeof cellValue === "string") {
      // 保留原有的单元格样式
      const originalFont = cell.font || {};
      const richText = [];
      let currentPos = 0;

      replacements.forEach((replacement) => {
        // 添加占位符前的文本（保持原格式）
        if (replacement.start > currentPos) {
          const beforeText = originalText.substring(
            currentPos,
            replacement.start
          );
          richText.push({
            text: beforeText,
            font: originalFont, // 保持原有格式
          });
        }

        // 添加替换后的文本（去掉加粗，但保持其他格式）
        richText.push({
          text: replacement.value,
          font: { ...originalFont, bold: false }, // 只去掉加粗
        });

        currentPos = replacement.end;
      });

      // 添加剩余文本
      if (currentPos < originalText.length) {
        const remainingText = originalText.substring(currentPos);
        richText.push({
          text: remainingText,
          font: originalFont, // 保持原有格式
        });
      }

      // 设置富文本
      cell.value = { richText: richText };
      return;
    }

    // 如果已经是富文本，需要更复杂的处理
    if (
      cellValue &&
      typeof cellValue === "object" &&
      "richText" in cellValue &&
      Array.isArray(cellValue.richText)
    ) {
      const newRichText = [];
      let textPosition = 0;

      // 遍历现有的富文本片段
      cellValue.richText.forEach((segment) => {
        const segmentText = segment.text || "";
        const segmentEnd = textPosition + segmentText.length;

        // 检查这个片段是否包含需要替换的占位符
        let hasReplacement = false;
        for (const replacement of replacements) {
          if (
            replacement.start < segmentEnd &&
            replacement.end > textPosition
          ) {
            hasReplacement = true;
            break;
          }
        }

        if (!hasReplacement) {
          // 如果没有占位符，直接保留原样
          newRichText.push(segment);
        } else {
          // 如果有占位符，需要分割处理
          let currentText = segmentText;
          let segmentPos = textPosition;

          // 找到在这个片段中的所有替换
          const segmentReplacements = replacements.filter(
            (r) => r.start >= textPosition && r.end <= segmentEnd
          );

          segmentReplacements.forEach((replacement) => {
            const relativeStart = replacement.start - segmentPos;
            const relativeEnd = replacement.end - segmentPos;

            // 添加占位符前的文本
            if (relativeStart > 0) {
              const beforeText = currentText.substring(0, relativeStart);
              newRichText.push({
                text: beforeText,
                font: segment.font, // 保持原有格式
              });
            }

            // 添加替换后的文本
            newRichText.push({
              text: replacement.value,
              font: { ...segment.font, bold: false }, // 只去掉加粗
            });

            // 更新当前文本和位置
            currentText = currentText.substring(relativeEnd);
            segmentPos = replacement.end;
          });

          // 添加剩余文本
          if (currentText.length > 0) {
            newRichText.push({
              text: currentText,
              font: segment.font, // 保持原有格式
            });
          }
        }

        textPosition = segmentEnd;
      });

      // 设置新的富文本
      cell.value = { richText: newRichText };
    }
  }

  /**
   * 插入图片到单元格
   * @private
   */
  async insertImageToCell(worksheet, rowNumber, colNumber, imageUrl) {
    try {
      // 下载图片
      const response = await axios.get(imageUrl, {
        responseType: "arraybuffer",
        timeout: 30000, // 30秒超时
        headers: {
          "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        },
      });

      // 获取图片扩展名（限制为 ExcelJS 支持的类型）
      const extension = this.getImageExtension(imageUrl);

      // 添加图片到工作簿
      const imageId = worksheet.workbook.addImage({
        buffer: response.data,
        extension: extension, // 现在类型匹配了
      });

      // 获取单元格位置
      const cell = worksheet.getCell(rowNumber, colNumber);

      // 设置行高和列宽以适应图片
      worksheet.getRow(rowNumber).height = 80; // 设置行高
      worksheet.getColumn(colNumber).width = 15; // 设置列宽

      // 插入图片
      worksheet.addImage(imageId, {
        tl: { col: colNumber - 1, row: rowNumber - 1 }, // 左上角位置
        ext: { width: 70, height: 105 }, // 图片大小
      });

      // 清空单元格内容（因为图片会覆盖）
      cell.value = "";

      console.log(`成功插入图片: ${imageUrl}`);
    } catch (error) {
      console.error(`图片插入失败: ${imageUrl}`, error.message);

      // 如果图片插入失败，使用链接文本
      const cell = worksheet.getCell(rowNumber, colNumber);
      cell.value = imageUrl;

      // 设置为超链接样式
      cell.font = {
        color: { argb: "FF0000FF" },
        underline: true,
      };
    }
  }

  /**
   * 获取图片扩展名 - 返回 ExcelJS 支持的类型
   * @private
   */
  getImageExtension(url) {
    const lowerUrl = url.toLowerCase();

    if (lowerUrl.includes(".png")) return "png";
    if (lowerUrl.includes(".jpg") || lowerUrl.includes(".jpeg")) return "jpeg";
    if (lowerUrl.includes(".gif")) return "gif";

    // 默认返回 png
    return "png";
  }

  /**
   * 判断是否为图片链接 - 只检查支持的格式
   * @private
   */
  isImageUrl(url) {
    // 检查是否为HTTP/HTTPS链接
    if (!/^https?:\/\//i.test(url)) {
      return false;
    }

    // 只检查 ExcelJS 支持的图片格式
    const supportedExtensions = [".jpg", ".jpeg", ".png", ".gif"];
    const lowerUrl = url.toLowerCase();

    return supportedExtensions.some((ext) => lowerUrl.includes(ext));
  }

  /**
   * 查找数组模板行
   * @private
   */
  findArrayTemplates(worksheet) {
    const templates = [];
    const processedRows = new Set(); // 记录已处理的行

    worksheet.eachRow((row, rowNumber) => {
      // 如果这一行已经处理过，跳过
      if (processedRows.has(rowNumber)) {
        return;
      }

      let foundArrayName = null;

      // 检查这一行是否有数组占位符
      row.eachCell((cell) => {
        if (cell.value && typeof cell.value === "string") {
          const match = cell.value.match(/\$\{(\w+):(\w+)\}/);
          if (match) {
            const [, arrayName] = match;
            if (!foundArrayName) {
              foundArrayName = arrayName;
            }
          }
        }
      });

      // 如果找到了数组占位符，添加到模板列表
      if (foundArrayName) {
        templates.push({
          row: rowNumber,
          arrayName: foundArrayName,
        });
        processedRows.add(rowNumber); // 标记为已处理
      }
    });

    return templates;
  }

  /**
   * 处理数组模板
   * @private
   */
  async processArrayTemplate(worksheet, template, data) {
    const { row, arrayName } = template;
    const arrayData = data[arrayName] || [];

    if (arrayData.length === 0) return;

    // 获取模板行
    const templateRow = worksheet.getRow(row);
    const templateValues = [];
    const templateStyles = [];

    // 保存模板行的值和样式
    templateRow.eachCell((cell, colNumber) => {
      templateValues[colNumber] = cell.value;
      templateStyles[colNumber] = {
        font: cell.font,
        alignment: cell.alignment,
        border: cell.border,
        fill: cell.fill,
        numFmt: cell.numFmt,
      };
    });

    // 如果数组长度大于1，需要插入额外的行
    if (arrayData.length > 1) {
      // 插入新行
      for (let i = 1; i < arrayData.length; i++) {
        worksheet.spliceRows(row + i, 0, []);
      }
    }

    // 填充所有数组数据
    for (let index = 0; index < arrayData.length; index++) {
      const item = arrayData[index];
      const targetRow = worksheet.getRow(row + index);

      // 填充数据和样式
      for (let colNumber = 1; colNumber < templateValues.length; colNumber++) {
        const value = templateValues[colNumber];

        if (value && typeof value === "string") {
          let newValue = value;

          // 替换数组字段占位符
          for (const key of Object.keys(item)) {
            const regex = new RegExp(`\\$\\{${arrayName}:${key}\\}`, "g");
            if (regex.test(newValue)) {
              const fieldValue = String(item[key]);

              // 检查是否为图片链接
              if (this.isImageUrl(fieldValue)) {
                await this.insertImageToCell(
                  worksheet,
                  row + index,
                  colNumber,
                  fieldValue
                );
                newValue = newValue.replace(regex, ""); // 清空占位符
              } else if (
                item[key] === null ||
                item[key] === undefined ||
                item[key] === " " ||
                item[key] === ""
              ) {
                console.log("值为空，替换占位符为空字符串");
                newValue = newValue.replace(regex, ""); // 替换为空字符串
              } else {
                console.log(`替换占位符: ${regex} -> ${fieldValue}`);
                // 即使不是图片链接，也要进行文本替换（包括空字符串）
                newValue = newValue.replace(regex, fieldValue);
              }
            }
          }

          const cell = targetRow.getCell(colNumber);
          cell.value = newValue;

          // 应用样式
          if (templateStyles[colNumber]) {
            if (templateStyles[colNumber].font) {
              cell.font = templateStyles[colNumber].font;
            }
            if (templateStyles[colNumber].alignment) {
              cell.alignment = templateStyles[colNumber].alignment;
            }
            if (templateStyles[colNumber].border) {
              cell.border = templateStyles[colNumber].border;
            }
            if (templateStyles[colNumber].fill) {
              cell.fill = templateStyles[colNumber].fill;
            }
            if (templateStyles[colNumber].numFmt) {
              cell.numFmt = templateStyles[colNumber].numFmt;
            }
          }
        }
      }
    }
  }

  /**
   * 判断是否为富文本对象
   * @private
   */
  isRichText(value) {
    return (
      value &&
      typeof value === "object" &&
      ("richText" in value || "text" in value)
    );
  }

  /**
   * 从富文本对象中提取纯文本
   * @private
   */
  extractTextFromRichText(richTextValue) {
    if (typeof richTextValue === "string") {
      return richTextValue;
    }

    if (
      richTextValue &&
      typeof richTextValue === "object" &&
      "richText" in richTextValue &&
      Array.isArray(richTextValue.richText)
    ) {
      return richTextValue.richText.map((item) => item.text || "").join("");
    }

    if (
      richTextValue &&
      typeof richTextValue === "object" &&
      "text" in richTextValue
    ) {
      return richTextValue.text;
    }

    return String(richTextValue);
  }
}

module.exports = ExcelService;
