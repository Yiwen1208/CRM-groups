//需求1:
// 第一个表格，一列是keyword(邮箱后缀)，一列 group
// 第二表格，一列是 email，一列是 group

//需求2:
//找不到keyword的email，出现两次的后缀，保留。

nodeXlsx = require("node-xlsx");
const nodeExcel = require("excel-export");
const fs = require("fs");

const obj = nodeXlsx?.parse("./data/data.xlsx");

const emails = obj[0].data;
let keywords = obj[1].data;

const needOne = () => {
  emails.map((email) => {
    let flag = false;
    if (email[0] === "Email") {
      return;
    }

    if (!email[0]) {
      return;
    }

    keywords.map((keyword) => {
      const keySub = email[0].split("@")[1];
      if (keySub.includes(keyword[0])) {
        email[1] = keyword[1];
        flag = true;
      }
    });
    if (!flag) {
      email[1] = "";
    }
  });
  const newEmail = [];
  emails.map((email) => {
    newEmail.push([email[0] ? email[0] : "", email[1] ? email[1] : ""]);
  });
  newEmail.shift();
  return newEmail;
};

const needTwo = () => {
  const newKeyword = [];

  const exceptEmails = {};
  let existKeywords = [];

  keywords.map((keyword) => {
    if (keyword[0] === "Keyword" || keyword[0] === undefined) {
      return;
    }
    existKeywords.push(keyword[0]);
  });
  existKeywords = Array.from(new Set(existKeywords));
  emails.map((email) => {
    if (email[0] === "Email") {
      return;
    }
    if (email[0] === undefined) {
      return;
    }
    const keySub = email[0].split("@")[1];
    if (existKeywords.includes(keySub)) {
      return;
    } else {
      if (exceptEmails[keySub]) {
        exceptEmails[keySub].push(email[0]);
      } else {
        exceptEmails[keySub] = [email[0]];
      }
    }
  });

  Object.keys(exceptEmails).map((key) => {
    if (exceptEmails[key].length > 1) {
      newKeyword.push(key);
    }
  });
  return newKeyword.map(keyword => [keyword]);
};

// excel配置
let conf = [
  {
    name: "match",
    cols: [
      {
        caption: "Email",
        type: "string",
      },
      {
        caption: "Group",
        type: "string",
      },
    ],
    rows: needOne()
  },
  {
    name: "withoutKeyword",
    cols: [
      {
        caption: "keyword",
        type: "string",
      },
    ],
    rows: needTwo()
  },
]; 

let result = nodeExcel.execute(conf);
let path = `${__dirname}/exportdata.xlsx`;
fs.writeFile(path, result, "binary", (err) => {
  err ? console.log(err) : null;
});
