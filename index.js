const fs = require("fs");
const xlsx = require("node-xlsx");
const { ocr: AipOcrClient } = require("baidu-aip-sdk");

// https://ai.baidu.com/ai-doc/OCR/rkibizxtw#%E5%A2%9E%E5%80%BC%E7%A8%8E%E5%8F%91%E7%A5%A8
// 设置APPID/AK/SK
const APP_ID = "24882769";
const API_KEY = "fAlcIN9XaPaTkiAy0YoYMGOZ";
const SECRET_KEY = "vHMHsZPAGLAbkWOlvecXzB1DDaDRAiEu";

const client = new AipOcrClient(APP_ID, API_KEY, SECRET_KEY);
const pdfDir = fs.readdirSync("./source-pdf");

const pdfNameJson = {};
let index = 0;

const recognitionPdf = () => {
  const pdfName = pdfDir[index];
  const pdfPath = `./source-pdf/${pdfName}`;
  console.log("index", index);
  console.log("pdfName", pdfName);
  client.vatInvoicePdf(pdfPath).then(function (result) {
    pdfNameJson[pdfName] = result;
    fs.renameSync(pdfPath, `./recognition-pdf/${pdfName}`);
  }).catch(function (err) {
    // 如果发生网络错误
    console.log(err);
  }).finally(() => {
    index = index + 1;
    if (index < pdfDir.length) {
      setTimeout(() => {
        recognitionPdf();
      }, 1000);
    } else {
      fs.writeFileSync("recognition-result.json", JSON.stringify(pdfNameJson, null, 2));
      generateXlsx();
    }
  });
};

const generateXlsx = () => {
  const invoiceFields = {
    InvoiceType: "发票种类",
    InvoiceTypeOrg: "发票名称",
    InvoiceCode: "发票代码",
    InvoiceNum: "发票号码",
    MachineNum: "机打号码",
    MachineCode: "机器编号",
    CheckCode: "校验码",
    InvoiceDate: "开票日期",
    PurchaserName: "购方名称",
    PurchaserRegisterNum: "购方纳税人识别号",
    PurchaserAddress: "购方地址及电话",
    PurchaserBank: "购方开户行及账号",
    Password: "密码区",
    Province: "省",
    City: "市",
    SheetNum: "联次",
    Agent: "是否代开",
    SellerName: "销售方名称",
    SellerRegisterNum: "销售方纳税人识别号",
    SellerAddress: "销售方地址及电话",
    SellerBank: "销售方开户行及账号",
    TotalAmount: "合计金额",
  };

  const pdfJSonMap = Object.keys(pdfNameJson).reduce((obj, pdfName) => {
    const invoiceInfo = pdfNameJson[pdfName];
    const { words_result } = invoiceInfo;
    const { InvoiceNum } = words_result;

    if (!obj[InvoiceNum]) {
      obj[InvoiceNum] = words_result;
      fs.renameSync(`./recognition-pdf/${pdfName}`, `./valid-pdf/${pdfName}`);
    }
    return obj;
  },{});

  const list = Object.values(pdfJSonMap).reduce((list, item, idx) => {
    Object.keys(invoiceFields).forEach((field) => {
      list[idx] = (list[idx] || []).concat(item[field]);
    });
    return list;
  }, []);

  list.unshift(Object.values(invoiceFields));

  const buffer = xlsx.build([{ data: list }]);
  fs.writeFileSync("result.xlsx", buffer);
};

recognitionPdf();