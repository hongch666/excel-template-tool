# excel-template-tool

一个强大的 Excel 模板工具，支持基于模板生成 Excel 文件，具备动态数据填充、富文本格式、数组数据和图片插入等功能。

## 特性

✨ **功能特性**

- 📋 基于 Excel 模板的动态数据填充
- 🎨 支持富文本格式保留（字体、颜色、粗体等）
- 📊 数组数据自动扩展和填充
- 🖼️ 图片 URL 自动下载并插入单元格
- 🔗 多占位符支持（基本字段 `${field}` 和数组字段 `${array:field}`）
- ⚡ 高性能，支持大数据量处理
- 🛡️ 错误处理和日志记录

## 安装

```bash
npm install excel-template-tool
```

或

```bash
yarn add excel-template-tool
```

## 占位符语法

### 基本占位符

在 Excel 模板中使用 `${fieldName}` 格式，将被数据对象中对应的值替换：

```
${supplierCompany}
${contactName}
${date}
```

### 数组占位符

对于需要重复填充的数组数据，使用 `${arrayName:fieldName}` 格式。工具会自动复制行来容纳所有数据：

```
${products:sku}        ${products:name}       ${products:totalAmount}
```

## 使用示例

### 示例 1：最小化示例

```javascript
const ExcelService = require("excel-template-tool");
const fs = require("fs");

const excelService = new ExcelService();

const data = {
  name: "John Doe",
  email: "john@example.com",
  amount: 1000,
};

(async () => {
  const buffer = await excelService.exportToExcel(data, "./template.xlsx");
  fs.writeFileSync("./output.xlsx", buffer);
  console.log("✅ Excel 文件已生成！");
})();
```

### 示例 2：发票模板（推荐）

```javascript
const ExcelService = require("excel-template-tool");
const fs = require("fs");
const path = require("path");

const excelService = new ExcelService();

const invoiceData = {
  // === 供应商信息 ===
  supplierCompany: "ABC Supplier Co., Ltd",
  supplierAddress: "123 Main Street, New York, NY 10001",
  supplierPhone: "+1-555-0123",
  supplierEmail: "supplier@abc.com",
  contactName: "John Doe",

  // === 买家信息 ===
  consignee: "XYZ Buyer Inc.",
  address: "456 Oak Avenue, Los Angeles, CA 90001",
  piNo: "PI-2024-0001",
  date: "2024-02-02",
  paymentTerms: "Net 30 Days",

  // === 产品列表（数组数据） ===
  products: [
    {
      sku: "PROD-001",
      selection: "Standard Selection",
      picture: "https://via.placeholder.com/100x150",
      totalAmount: 1000,
    },
    {
      sku: "PROD-002",
      selection: "Premium Selection",
      picture: "https://via.placeholder.com/100x150",
      totalAmount: 2000,
    },
    {
      sku: "PROD-003",
      selection: "Deluxe Selection",
      picture: "https://via.placeholder.com/100x150",
      totalAmount: 1500,
    },
  ],

  // === 汇总 ===
  totalAmount: 4500,
  totalPrice: 4500,
  totalShipping: 200,
};

(async () => {
  try {
    console.log("🚀 正在生成 Excel 文件...");

    const buffer = await excelService.exportToExcel(
      invoiceData,
      "./templates/invoice-template.xlsx"
    );

    fs.writeFileSync("./output/invoice.xlsx", buffer);
    console.log("✅ 发票已生成：invoice.xlsx");
    console.log(`📊 产品数量：${invoiceData.products.length}`);
    console.log(`💰 总金额：$${invoiceData.totalAmount}`);
  } catch (error) {
    console.error("❌ 生成失败:", error.message);
  }
})();
```

### 示例 3：Express.js 中使用

```javascript
const express = require("express");
const ExcelService = require("excel-template-tool");

const app = express();
const excelService = new ExcelService();

app.post("/api/export-invoice", async (req, res) => {
  try {
    const invoiceData = req.body;

    const buffer = await excelService.exportToExcel(
      invoiceData,
      "./templates/invoice-template.xlsx"
    );

    // 设置响应头，下载文件
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="invoice.xlsx"');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(3000, () => {
  console.log("Server running on port 3000");
});
```

### 示例 4：图片自动插入

如果数据中的字段值是 HTTP/HTTPS 图片链接，工具会自动下载并插入到 Excel 单元格中：

```javascript
const data = {
  products: [
    {
      sku: "PROD-001",
      name: "Product A",
      // 自动下载并插入到单元格
      picture: "https://example.com/product-a.jpg",
    },
    {
      sku: "PROD-002",
      name: "Product B",
      picture: "https://example.com/product-b.png",
    },
  ],
};

// 支持的图片格式：JPG, JPEG, PNG, GIF
```

## API 文档

### ExcelService

#### `exportToExcel(data, templatePath)`

生成 Excel 文件 Buffer。

**参数：**

- `data` (Object): 要填充的数据对象
- `templatePath` (String): 模板文件路径（绝对路径或相对于工作目录的相对路径）

**返回值：** `Promise<Buffer>` - Excel 文件的 Buffer

**异常：**

- 如果模板文件不存在会抛出错误
- 如果工作表不存在会抛出错误

**示例：**

```javascript
const buffer = await excelService.exportToExcel(data, "./template.xlsx");
```

## 模板文件要求

### 模板结构

1. **基本信息字段**: 使用 `${fieldName}` 格式
2. **数组数据**: 在单独的行中使用 `${arrayName:fieldName}` 格式
   - 一行只能对应一个数组
   - 工具会自动复制行来容纳数组数据

### 示例模板结构

```
第1行:   [标题] Proforma Invoice
第2行:   供应商信息
第3行:   供应商: ${supplierCompany}     地址: ${supplierAddress}
第4行:   联系人: ${contactName}         邮箱: ${supplierEmail}
...
第12行:  SKU              产品名称              数量       单价        总额
第13行:  ${products:sku}  ${products:name}    ${products:qty}  ${products:price}  ${products:totalAmount}
...
第20行:  总金额: ${totalAmount}
```

## 配置选项

默认配置（可在源代码中修改）：

- 图片超时时间: 30 秒
- 图片大小: 70x105px
- 行高: 80
- 列宽: 15

## 特性详解

### 1. 富文本格式保留

工具会自动保留原 Excel 模板中的格式化信息：

- 字体样式
- 颜色
- 粗体/斜体
- 对齐方式
- 边框
- 背景色

### 2. 图片自动下载

支持自动下载并插入 HTTP/HTTPS 图片链接，支持格式：JPG, PNG, GIF

图片下载失败时会自动降级为蓝色超链接文本。

### 3. 错误处理

所有错误都会被捕获并记录到控制台：

```
图片插入失败: https://example.com/image.jpg
```

## 注意事项

⚠️ **重要提示**

1. 确保模板文件存在且可读
2. 图片下载需要网络连接
3. 数组字段应为数组类型，其他字段应为基本类型
4. 同一行中只能有一个数组占位符
5. 空值（null、undefined、''）会被替换为空字符串

## 许可证

MIT

## 常见问题

**Q: 如何处理空值？**
A: 如果数据字段为 null、undefined 或空字符串，占位符会被替换为空字符串。

**Q: 是否支持多个工作表？**
A: 目前只支持第一个工作表。

**Q: 图片下载超时怎么办？**
A: 超时时间设置为 30 秒，超时的图片会降级为链接文本。

**Q: 如何自定义图片大小？**
A: 在 ExcelService.js 中修改 `insertImageToCell` 方法中的 `ext` 参数。

## 更新日志

### v1.0.0 (2024-02-02)

- 🎉 初版发布
- ✨ 支持基本占位符和数组占位符
- 🖼️ 支持图片自动插入
- 🎨 支持富文本格式保留

## 贡献

欢迎提交 Issue 和 Pull Request！

---

**Made with ❤️ for Excel templates**
