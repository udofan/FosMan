using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static FosMan.Enums;

namespace FosMan {
    /// <summary>
    /// Тип парсинга
    /// </summary>
    public enum EParseType {
        /// <summary>
        /// В текущей строке
        /// </summary>
        Inline,
        /// <summary>
        /// Захват строк
        /// </summary>
        Multiline
    }

    /// <summary>
    /// Парсер документа
    /// </summary>
    internal static class DocParser {
        /// <summary>
        /// Парсинг docx-файла
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="targetObj">целевой объект</param>
        /// <param name="rules"></param>
        public static bool TryParse<T>(string fileName, T targetObj, IEnumerable<IDocParseRule<T>> rules, out ErrorList errors) where T : BaseObj {
            var result = false;
            errors = new();

            try {
                //активация правил
                var activeRules = rules?.ToHashSet();
                //foreach (var rule in rules) rule.Disabled = false;

                if (File.Exists(fileName)) {
                    using (var docx = DocX.Load(fileName)) {

                        IDocParseRule<T> catchingRule = null;
                        StringBuilder catchingValue = new();
                        
                        //цикл по параграфам
                        foreach (var par in docx.Paragraphs) {
                            var text = par.Text.Trim(' ', '\r', '\n', '\t');

                            if (catchingRule != null) { //если в режиме отлова строк для правила catchingRule
                                if (catchingRule.StopMarkers == null) {
                                    //errors.Add($"Для правила {catchingRule.PropertyName} с типом {catchingRule.Type} не заданы StopMarkers");
                                    errors.Add(EErrorType.ParseRuleMissingStopMarkers, $"правило {catchingRule.PropertyName}");
                                    catchingRule = null; //сброс захвата
                                    continue;
                                }
                                //var stopCatch = catchingRule.StopMarkers.Any(r => r.IsMatch(text));
                                (Regex regex, Match match, int idx) m = catchingRule.StopMarkers.Select(x => (x.marker, x.marker.Match(text), x.catchGroupIdx)).FirstOrDefault(x => x.Item2.Success);
                                if (m.match != null) {
                                    if (m.idx >= 0) { //требуется захват значения, где сработал маркер останова
                                        if (!string.IsNullOrEmpty(catchingRule.MultilineConcatValue) && catchingValue.Length > 0) {
                                            catchingValue.Append(catchingRule.MultilineConcatValue);
                                        }
                                        catchingValue.Append(m.match.Groups[m.idx].Value);
                                    }
                                    //захват будем считать завершенным, когда было захвачено значение
                                    if (catchingValue.Length > 0) {
                                        ApplyValue(catchingRule, null, targetObj, text, catchingValue.ToString(), par, errors);
                                        catchingValue.Clear();
                                        catchingRule = null; //сброс правила-ловца строк
                                    }
                                }
                                else { //продолжаем захват
                                    if (!string.IsNullOrEmpty(catchingRule.MultilineConcatValue) && catchingValue.Length > 0) {
                                        catchingValue.Append(catchingRule.MultilineConcatValue);
                                    }
                                    catchingValue.Append(text);
                                }
                            }
                            else { //отлова нет
                                if (string.IsNullOrWhiteSpace(text)) {
                                    continue; //пустые строки пропускаем
                                }
                                foreach (var rule in activeRules.ToHashSet()) {
                                    (Regex regex, Match match, int idx) m = rule.StartMarkers.Select(x => (x.marker, x.marker.Match(text), x.catchGroupIdx)).FirstOrDefault(x => x.Item2.Success);
                                    if (m.match != null) {
                                        if (!rule.MultyApply) {
                                            activeRules.Remove(rule);   //правило сработало - отключаем
                                        }
                                        //инлайн поиск значения
                                        if (rule.Type == EParseType.Inline) {
                                            ApplyValue(rule, m, targetObj, text, null, par, errors);
                                        }
                                        else if (rule.Type == EParseType.Multiline) {
                                            catchingValue.Clear();
                                            if (m.idx >= 0) { //требуется захват с маркера начала
                                                catchingValue.Append(m.match.Groups[m.idx].Value);
                                            }
                                            catchingRule = rule;
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                result = true;
            }
            catch (Exception ex) {
                errors.AddException(ex);
                //errors?.Add($"{ex.Message}\r\n{ex.StackTrace}");
                var tt = 0;
            }

            return result;
        }

        /// <summary>
        /// Применить правило
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rule"></param>
        /// <param name="text"></param>
        /// <param name="targetObj"></param>
        static public void ApplyValue<T>(IDocParseRule<T> rule,
                                         (Regex regex, Match match, int groupIdx)? ruleMatch,
                                         T targetObj,
                                         string text,
                                         string value,
                                         Paragraph par,
                                         ErrorList errors) where T : BaseObj {
            if (!string.IsNullOrEmpty(rule.PropertyName)) {
                if (ruleMatch.HasValue) {
                    if (ruleMatch.Value.groupIdx < ruleMatch.Value.match.Groups.Count) {
                        value = ruleMatch.Value.match.Groups[ruleMatch.Value.groupIdx].Value;
                    }
                    else {
                        //errors.Add($"Для правила {rule.PropertyName} с типом {rule.Type} задан некорректный [inlineGroupIdx] " +
                        //           $"для StartMarker [{ruleMatch.Value.regex}]");
                        errors.Add(EErrorType.ParseRuleWrongInlineGroupIdxInStartMarkers, $"правило {rule.PropertyName}, значение StartMarker: {ruleMatch.Value.regex}");
                        return;
                    }
                }
                if (rule.TrimChars != null) {
                    value = value.Trim(rule.TrimChars);
                }
                targetObj.SetProperty(rule.PropertyName, null, value);
            }
            else if (rule.Action != null) {
                var args = new DocParseRuleActionArgs<T>() {
                    Target = targetObj,
                    Rule = rule,
                    Match = ruleMatch.Value.match,
                    Text = text,
                    Value = value,
                    Paragraph = par
                };
                rule.Action.Invoke(args);
            }
        }
    }
}
