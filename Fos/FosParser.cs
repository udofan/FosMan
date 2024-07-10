using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Парсер ФОСов
    /// </summary>
    internal static class FosParser {
        static List<IDocParseRule<Fos>> m_rules = new() {
            new FosParseRuleYear(), 
            new FosParseRuleCompiler(), 
            new FosParseRuleDepartment(), 
            new FosParseRuleDisciplineName(),
            new FosParseRuleDirection(),
            new FosParseRuleFormsOfStudy(),
            new FosParseRuleProfileInline(),
            new FosParseRuleProfileMultiline(),
            new FosParseRuleEvaluationTools()
        };

        public static List<IDocParseRule<Fos>> Rules { get => m_rules; }
    }
}
