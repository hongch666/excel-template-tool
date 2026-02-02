/**
 * excel-template-tool 完整使用示例
 */

const ExcelService = require("./lib/ExcelService");
const path = require("path");
const fs = require("fs");

async function example() {
  const excelService = new ExcelService();

  // 示例数据 - 根据您的模板调整
  const data = {
    // === 供应商信息 ===
    supplierCompany: "ABC Supplier Co., Ltd",
    supplierAddress: "123 Main Street, New York, NY 10001",
    supplierPhone: "+1-555-0123",
    supplierEmail: "supplier@abc.com",

    // === 联系人信息 ===
    contactName: "John Doe",
    contactEmail: "john.doe@abc.com",

    // === 买家信息 ===
    consignee: "XYZ Buyer Inc.",
    address: "456 Oak Avenue, Los Angeles, CA 90001",
    name: "Product Name",
    piNo: "PI-2024-0001",
    shopmentTerms: "FOB Shanghai",
    date: "2024-02-02",
    paymentTerms: "Net 30 Days",

    // === 产品列表 ===
    products: [
      {
        picture: "https://via.placeholder.com/100x150?text=Product+1",
        selection: "Standard Selection",
        totalAmount: 1000,
        sku: "SKU001",
      },
      {
        picture: "https://via.placeholder.com/100x150?text=Product+2",
        selection: "Premium Selection",
        totalAmount: 2000,
        sku: "SKU002",
      },
      {
        picture: "https://via.placeholder.com/100x150?text=Product+3",
        selection: "Deluxe Selection",
        totalAmount: 1500,
        sku: "SKU003",
      },
    ],

    // === 汇总信息 ===
    totalAmount: 4500,
    totalPrice: 4500,
    totalShipping: 200,
  };

  try {
    console.log("🚀 开始生成 Excel 文件...\n");

    // 模板路径
    const templatePath = path.join(__dirname, "./PI-template.xlsx");

    // 检查模板文件是否存在
    if (!fs.existsSync(templatePath)) {
      console.error("❌ 模板文件不存在:", templatePath);
      console.error("请确保 PI-template.xlsx 文件在项目根目录中");
      return;
    }

    // 生成 Excel 文件
    const buffer = await excelService.exportToExcel(data, templatePath);

    // 保存文件
    const outputPath = path.join(__dirname, "./output.xlsx");
    fs.writeFileSync(outputPath, buffer);

    console.log("✅ Excel 文件已成功生成！");
    console.log(`📁 文件位置: ${outputPath}`);
    console.log(`📊 文件大小: ${(buffer.length / 1024).toFixed(2)} KB\n`);

    // 打印数据摘要
    console.log("📋 数据摘要:");
    console.log(`   供应商: ${data.supplierCompany}`);
    console.log(`   买家: ${data.consignee}`);
    console.log(`   产品数量: ${data.products.length}`);
    console.log(`   总金额: $${data.totalAmount}`);
  } catch (error) {
    console.error("❌ 生成失败:", error.message);
    console.error(error.stack);
  }
}

// 运行示例
if (require.main === module) {
  example();
}

module.exports = { example };
