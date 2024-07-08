using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Парсер РПД
    /// </summary>
    internal static class RpdParser {
        static List<IDocParseRule<Rpd>> m_rules = new() {
            //new RpdParseRuleYear(), 
            new RpdParseRuleCompilerInline(),
            new RpdParseRuleCompilerMultiline(),
            //new RpdParseRuleDepartment(), 
            //new RpdParseRuleDisciplineName(),
            //new RpdParseRuleDirection(),
            //new RpdParseRuleFormsOfStudyInline(),
            //new RpdParseRuleFormsOfStudyMultiline(),
            //new RpdParseRuleProfileInline(),
            //new RpdParseRuleProfileMultiline(),
            new RpdParseRulePrevDisciplines(),
            new RpdParseRuleNextDisciplines(),
            new RpdParseRulePurpose(),
            //new RpdParseRuleTasks(),
            new RpdParseRuleDescription()
        };

        public static List<IDocParseRule<Rpd>> Rules { get => m_rules; }
    }
}
