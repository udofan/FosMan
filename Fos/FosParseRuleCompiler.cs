﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class FosParseRuleCompiler : IDocParseRule<Fos> {
        //public bool Disabled { get; set; }
        public EParseType Type { get; set; } = EParseType.Inline;
        public string MultilineConcatValue { get; set; } = string.Empty;
        public string PropertyName { get; set; } = nameof(Fos.Compiler);
        public Type PropertyType { get; set; } = typeof(Fos).GetProperty(nameof(Fos.Compiler))?.PropertyType;
        public List<(Regex marker, int catchGroupIdx)> StartMarkers { get; set; } = [
            (new(@"([А-Я]{1}\.\s*[А-Я]{1}\.\s*.+$|[А-Яа-я]{2,}\s+[А-Я]{1}\.\s*[А-Я]{1}\.)", RegexOptions.Compiled | RegexOptions.IgnoreCase), 1)
        ];
        public List<(Regex marker, int catchGroupIdx)> StopMarkers { get; set; } = null;
        public char[] TrimChars { get; set; } = null;
        public Action<DocParseRuleActionArgs<Fos>> Action { get; set; } = null;
        public bool MultyApply { get; set; } = false;
        //public bool Equals(IDocParseRule<Fos>? other) {
        //    return string.Compare(this.Name, other?.Name) == 0;
        //}

        public bool Equals<T>(IDocParseRule<T>? other) {
            throw new NotImplementedException();
        }
    }
}
