// Office Scripts rules that the TypeScript compiler does not check, using Microsoft's official ExcelScript types
// (test/excelscript-official.d.ts, from OfficeDev/office-scripts-docs-reference):
//   1. array method callbacks must be inline arrow functions;
//   2. no explicit 'any';
//   3. no aliasing / storing of Office Scripts API objects: an Excel object (Worksheet, Table, Range, ...) may only be
//      held in a variable filled straight from an Excel call (or a function parameter) - never stored in an object or
//      array literal, written to a property, or copied from another variable or property.
// Usage: node lint_officescript.mjs ../asset_audit_officescript.ts [more files]
import ts from "typescript";
import path from "node:path";
import { fileURLToPath } from "node:url";
const here = path.dirname(fileURLToPath(import.meta.url));
const METHODS = new Set(["map", "forEach", "filter", "sort", "reduce", "reduceRight", "find", "findIndex", "findLast",
  "findLastIndex", "some", "every", "flatMap"]);
const files = process.argv.slice(2).map(f => path.resolve(f));
const program = ts.createProgram(files.concat([path.join(here, "excelscript-official.d.ts")]),
  { strict: true, target: ts.ScriptTarget.ES2017, lib: ["lib.es2017.d.ts"], noEmit: true });
const checker = program.getTypeChecker();
let problems = 0;
// True when the type is (or contains, e.g. a union or an array of) an ExcelScript interface type
function isApiType(t, depth = 0) {
  if (!t || depth > 4) return false;
  if (t.isUnion && t.isUnion()) return t.types.some(x => isApiType(x, depth + 1));
  const sym = t.getSymbol() || t.aliasSymbol;
  if (sym) {
    const name = checker.getFullyQualifiedName(sym);
    if (/^ExcelScript\.[A-Z]/.test(name) && !(t.flags & ts.TypeFlags.EnumLike) && !(sym.flags & ts.SymbolFlags.EnumMember) &&
      !(sym.flags & ts.SymbolFlags.Enum)) return true;
  }
  if (checker.isArrayType && checker.isArrayType(t)) return isApiType(checker.getTypeArguments(t)[0], depth + 1);
  return false;
}
// Allowed sources of an Excel object: a call (or ?? / || / ?: of calls), an array literal of calls is NOT allowed
function fromCall(e) {
  while (ts.isParenthesizedExpression(e) || ts.isAsExpression(e) || ts.isNonNullExpression(e)) e = e.expression;
  if (ts.isCallExpression(e)) return true;
  if (ts.isConditionalExpression(e)) return fromCall(e.whenTrue) && fromCall(e.whenFalse);
  if (ts.isBinaryExpression(e) && [ts.SyntaxKind.BarBarToken, ts.SyntaxKind.QuestionQuestionToken].includes(e.operatorToken.kind))
    return fromCall(e.left) && fromCall(e.right);
  return e.kind === ts.SyntaxKind.NullKeyword;
}
for (const file of files) {
  const src = program.getSourceFile(file);
  const report = (node, msg) => {
    const { line, character } = src.getLineAndCharacterOfPosition(node.getStart());
    console.log(`${path.relative(process.cwd(), file)}:${line + 1}:${character + 1} ${msg}`); problems++;
  };
  const api = e => { try { return isApiType(checker.getTypeAtLocation(e)); } catch { return false; } };
  const visit = node => {
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) && METHODS.has(node.expression.name.text)) {
      const cb = node.arguments[0];
      if (cb && !ts.isArrowFunction(cb)) report(cb, `'${node.expression.name.text}' callback is not an inline arrow function`);
    }
    if (node.kind === ts.SyntaxKind.AnyKeyword) report(node, "explicit 'any'");
    if (ts.isVariableDeclaration(node) && node.initializer && api(node.initializer) && !fromCall(node.initializer))
      report(node.initializer, "Excel object copied into a variable (not from an Excel call)");
    if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.EqualsToken && api(node.right)) {
      if (!ts.isIdentifier(node.left)) report(node, "Excel object written to a property");
      else if (!fromCall(node.right)) report(node.right, "Excel object copied into a variable (not from an Excel call)");
    }
    if ((ts.isPropertyAssignment(node) || ts.isShorthandPropertyAssignment(node)) &&
      api(ts.isPropertyAssignment(node) ? node.initializer : node.name)) report(node, "Excel object stored in an object literal");
    if (ts.isArrayLiteralExpression(node) && node.elements.some(e => api(e))) report(node, "Excel object stored in an array literal");
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) &&
      ["push", "unshift", "concat", "set"].includes(node.expression.name.text) && node.arguments.some(a => api(a)) &&
      !api(node.expression.expression))
      report(node, `Excel object stored with '${node.expression.name.text}'`);
    ts.forEachChild(node, visit);
  };
  visit(src);
}
console.log(problems ? `${problems} problem(s)` : "Office Scripts lint: no problems");
process.exit(problems ? 1 : 0);
