// Functions people use in Google Sheets and Excel that HyperFormula 3.4 does not have. Registered once,
// on the server and in the browser, so both work them out the same way.
import { HyperFormula, FunctionPlugin, FunctionArgumentType, CellError, ErrorType } from 'hyperformula';

const ARG = FunctionArgumentType;
const num = (v) => (typeof v === 'number' ? v : v && typeof v === 'object' && typeof v.val === 'number' ? v.val : null);
const flat = (range) => (range && typeof range.valuesFromTopLeftCorner === 'function' ? range.valuesFromTopLeftCorner() : [range]);

class TextExtras extends FunctionPlugin {
  concat(ast, state) {
    return this.runFunction(ast.args, state, this.metadata('CONCAT'), (...args) => ''.concat(...args));
  }
  regexmatch(ast, state) {
    return this.runFunction(ast.args, state, this.metadata('REGEXMATCH'), (text, re) => {
      try { return new RegExp(re).test(text); } catch { return new CellError(ErrorType.VALUE, 'Not a valid regular expression.'); }
    });
  }
  regexextract(ast, state) {
    return this.runFunction(ast.args, state, this.metadata('REGEXEXTRACT'), (text, re) => {
      let m;
      try { m = new RegExp(re).exec(text); } catch { return new CellError(ErrorType.VALUE, 'Not a valid regular expression.'); }
      if (!m) return new CellError(ErrorType.NA, 'No match.');
      return m.length > 1 ? m[1] ?? '' : m[0];
    });
  }
  regexreplace(ast, state) {
    return this.runFunction(ast.args, state, this.metadata('REGEXREPLACE'), (text, re, by) => {
      try { return text.replace(new RegExp(re, 'g'), by); } catch { return new CellError(ErrorType.VALUE, 'Not a valid regular expression.'); }
    });
  }
}
TextExtras.implementedFunctions = {
  CONCAT: { method: 'concat', parameters: [{ argumentType: ARG.STRING }], repeatLastArgs: 1, expandRanges: true },
  REGEXMATCH: { method: 'regexmatch', parameters: [{ argumentType: ARG.STRING }, { argumentType: ARG.STRING }] },
  REGEXEXTRACT: { method: 'regexextract', parameters: [{ argumentType: ARG.STRING }, { argumentType: ARG.STRING }] },
  REGEXREPLACE: { method: 'regexreplace', parameters: [{ argumentType: ARG.STRING }, { argumentType: ARG.STRING }, { argumentType: ARG.STRING }] },
};

class StatExtras extends FunctionPlugin {
  // RANK(number, range, [ascending]): 1 for the largest (or the smallest when ascending), ties share a rank.
  rank(ast, state) {
    return this.runFunction(ast.args, state, this.metadata('RANK'), (n, range, asc) => {
      const vals = flat(range).map(num).filter((v) => v !== null);
      if (!vals.includes(n)) return new CellError(ErrorType.NA, 'The number is not in the range.');
      return 1 + vals.filter((v) => (asc ? v < n : v > n)).length;
    });
  }
  rankavg(ast, state) {
    return this.runFunction(ast.args, state, this.metadata('RANK.AVG'), (n, range, asc) => {
      const vals = flat(range).map(num).filter((v) => v !== null);
      const ties = vals.filter((v) => v === n).length;
      if (!ties) return new CellError(ErrorType.NA, 'The number is not in the range.');
      const before = vals.filter((v) => (asc ? v < n : v > n)).length;
      return before + (ties + 1) / 2;
    });
  }
}
StatExtras.implementedFunctions = {
  RANK: { method: 'rank', parameters: [{ argumentType: ARG.NUMBER }, { argumentType: ARG.RANGE }, { argumentType: ARG.BOOLEAN, defaultValue: false }] },
  'RANK.EQ': { method: 'rank', parameters: [{ argumentType: ARG.NUMBER }, { argumentType: ARG.RANGE }, { argumentType: ARG.BOOLEAN, defaultValue: false }] },
  'RANK.AVG': { method: 'rankavg', parameters: [{ argumentType: ARG.NUMBER }, { argumentType: ARG.RANGE }, { argumentType: ARG.BOOLEAN, defaultValue: false }] },
};

// AVERAGEIFS reuses HyperFormula's own SUMIFS machinery (criteria, caching), so it matches SUMIFS / COUNTIFS.
function averageIfsPlugin() {
  const Base = HyperFormula.getFunctionPlugin('SUMIFS');
  class AverageIfs extends Base {
    averageifs(ast, state) {
      const name = 'AVERAGEIFS';
      return this.runFunction(ast.args, state, this.metadata(name), (values, ...args) => {
        const r = this.computeConditionalAggregationFunction(values, args, name, { s: 0, n: 0 },
          (a, b) => ({ s: a.s + b.s, n: a.n + b.n }),
          (v) => (typeof num(v) === 'number' && typeof v !== 'boolean' ? { s: num(v), n: 1 } : { s: 0, n: 0 }));
        if (r instanceof CellError) return r;
        return r.n ? r.s / r.n : new CellError(ErrorType.DIV_BY_ZERO);
      });
    }
  }
  AverageIfs.implementedFunctions = { AVERAGEIFS: { method: 'averageifs', parameters: [{ argumentType: ARG.RANGE }, { argumentType: ARG.RANGE }, { argumentType: ARG.NOERROR }], repeatLastArgs: 2 } };
  return AverageIfs;
}

const names = (plugin) => Object.fromEntries(Object.keys(plugin.implementedFunctions).map((k) => [k, k]));
let done = false;
export function registerFunctions() {
  if (done) return;
  done = true;
  const avg = averageIfsPlugin();
  for (const p of [TextExtras, StatExtras, avg]) {
    const tr = names(p);
    HyperFormula.registerFunctionPlugin(p, { enGB: tr, enUS: tr });
  }
}

export const EXTRA_FUNCTIONS = ['CONCAT', 'REGEXMATCH', 'REGEXEXTRACT', 'REGEXREPLACE', 'RANK', 'RANK.EQ', 'RANK.AVG', 'AVERAGEIFS'];
