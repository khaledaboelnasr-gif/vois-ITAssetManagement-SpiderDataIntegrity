// Office Scripts rules the TypeScript compiler does not check: array method callbacks must be inline arrow functions,
// and no explicit 'any'. Usage: node lint_officescript.mjs ../asset_audit_officescript.ts [more files]
import ts from "typescript";
import fs from "node:fs";
const METHODS = new Set(["map", "forEach", "filter", "sort", "reduce", "reduceRight", "find", "findIndex", "findLast",
  "findLastIndex", "some", "every", "flatMap"]);
let problems = 0;
for (const file of process.argv.slice(2)) {
  const src = ts.createSourceFile(file, fs.readFileSync(file, "utf8"), ts.ScriptTarget.ES2017, true);
  const report = (node, msg) => {
    const { line, character } = src.getLineAndCharacterOfPosition(node.getStart());
    console.log(`${file}:${line + 1}:${character + 1} ${msg}`); problems++;
  };
  const visit = node => {
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) && METHODS.has(node.expression.name.text)) {
      const cb = node.arguments[0];
      if (cb && !ts.isArrowFunction(cb)) report(cb, `'${node.expression.name.text}' callback is not an inline arrow function`);
    }
    if (node.kind === ts.SyntaxKind.AnyKeyword) report(node, "explicit 'any'");
    ts.forEachChild(node, visit);
  };
  visit(src);
}
console.log(problems ? `${problems} problem(s)` : "Office Scripts lint: no problems");
process.exit(problems ? 1 : 0);
