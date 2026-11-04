import { readFileSync, writeFileSync } from "node:fs";
import { parseTemplate } from "./packages/template-parser/dist/index.js";
const t = await parseTemplate(readFileSync("tests/fixtures/templates/acm-journal.docx"), { documentType: "journal", organization: "ACM" });
writeFileSync("E:/各类文件/DSH 工作文件夹/_可视化验证/acm-template.json", JSON.stringify(t, null, 2), "utf8");
console.log("page:", JSON.stringify(t.page && { w: t.page.width, h: t.page.height }));
console.log("requirements:", JSON.stringify(t.requirements));
console.log("warnings:", (t.warnings ?? []).length);
