using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xceed.Document.NET;

namespace FosMan {
    internal static class DocXEx {
        static ConcurrentDictionary<Document, ConcurrentDictionary<Paragraph, (List list, int parIndex)>> m_cacheListParagraphs = [];
        static ConcurrentDictionary<Document, ConcurrentDictionary<int, ConcurrentDictionary<int, Paragraph>>> m_cacheListLevels = [];
        //static ConcurrentDictionary<Document, ConcurrentDictionary<int, List>> m_cacheLists = [];

        static public Stopwatch SwBase = new();
        static public Stopwatch SwFast = new();

        internal static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /*
        internal static int GetNumId(this Paragraph par) {
            XElement xElement = par.Xml.Descendants().FirstOrDefault((XElement s) => s.Name.LocalName == "numId");
            if (xElement == null) {
                return -1;
            }

            return int.Parse(xElement.Attribute(w + "val").Value);
        }

        internal static int GetListItemLevel(this Paragraph par) {
            if (par.ParagraphNumberProperties != null) {
                XElement xElement = par.ParagraphNumberProperties.Descendants().FirstOrDefault((XElement el) => el.Name.LocalName == "ilvl");
                if (xElement != null) {
                    return int.Parse(xElement.Attribute(w + "val").Value);
                }
            }

            return -1;
        }

        internal static string GetAbstractNumIdValue(Document document, string numId) {
            if (document == null)
                return null;
            if (numId == null)
                return null;

            // Find num node in numbering.
            var documentNumberingDescendants = document._numbering.Descendants();
            var numNodes = documentNumberingDescendants.Where(n => n.Name.LocalName == "num");
            var numNode = numNodes.FirstOrDefault(node => node.Attribute(w + "numId").Value.Equals(numId));
            if (numNode == null)
                return null;

            // Get abstractNumId node and its value from numNode.
            var abstractNumIdNode = numNode.Descendants().FirstOrDefault(n => n.Name.LocalName == "abstractNumId");
            if (abstractNumIdNode == null)
                return null;
            var abstractNumNodeValue = abstractNumIdNode.Attribute(w + "val").Value;
            if (string.IsNullOrEmpty(abstractNumNodeValue))
                return null;

            return abstractNumNodeValue;
        }

        internal static XElement GetAbstractNum(Document document, string numId) {
            if (document == null)
                return null;
            if (numId == null)
                return null;

            var abstractNumNodeValue = GetAbstractNumIdValue(document, numId);
            if (string.IsNullOrEmpty(abstractNumNodeValue))
                return null;

            // Find abstractNum node in numbering.
            var documentNumberingDescendants = document._numbering.Descendants();
            var abstractNumNodes = documentNumberingDescendants.Where(n => n.Name.LocalName == "abstractNum");
            var abstractNumNode = abstractNumNodes.FirstOrDefault(node => node.Attribute(w + "abstractNumId").Value.Equals(abstractNumNodeValue));

            return abstractNumNode;
        }

        internal static XElement GetStyle(Document fileToConvert, string styleId) {
            if (fileToConvert == null)
                throw new ArgumentNullException("fileToConvert");
            if (string.IsNullOrEmpty(styleId))
                throw new ArgumentNullException("styleId");

            var styles = fileToConvert._styles.Element(XName.Get("styles", w.NamespaceName));
            return styles.Elements(XName.Get("style", w.NamespaceName))
                         .FirstOrDefault(x => (x.Attribute(XName.Get("styleId", w.NamespaceName)) != null) && (x.Attribute(XName.Get("styleId", w.NamespaceName)).Value == styleId));
        }

        private static int GetLinkedStyleNumId(Document document, string numId) {
            Debug.Assert(document != null, "document should not be null");

            var abstractNumElement = GetAbstractNum(document, numId);
            if (abstractNumElement != null) {
                var numStyleLink = abstractNumElement.Element(XName.Get("numStyleLink", w.NamespaceName));
                if (numStyleLink != null) {
                    var val = numStyleLink.Attribute(XName.Get("val", w.NamespaceName));
                    if (!string.IsNullOrEmpty(val.Value)) {
                        var linkedStyle = GetStyle(document, val.Value);
                        if (linkedStyle != null) {
                            var linkedNumId = linkedStyle.Descendants(XName.Get("numId", w.NamespaceName)).FirstOrDefault();
                            if (linkedNumId != null) {
                                var linkedNumIdVal = linkedNumId.Attribute(XName.Get("val", w.NamespaceName));
                                if (!string.IsNullOrEmpty(linkedNumIdVal.Value))
                                    return Int32.Parse(linkedNumIdVal.Value);
                            }
                        }
                    }
                }
            }

            return -1;
        }

        internal static string GetListItemType(Paragraph p, Document document) {
            var paragraphNumberPropertiesDescendants = p.ParagraphNumberProperties.Descendants();
            var ilvlNode = paragraphNumberPropertiesDescendants.FirstOrDefault(el => el.Name.LocalName == "ilvl");
            var ilvlValue = (ilvlNode != null) ? ilvlNode.Attribute(w + "val").Value : null;

            var numIdNode = paragraphNumberPropertiesDescendants.FirstOrDefault(el => el.Name.LocalName == "numId");
            var numIdValue = (numIdNode != null) ? numIdNode.Attribute(w + "val").Value : null;

            var abstractNumNode = GetAbstractNum(document, numIdValue);
            if (abstractNumNode != null) {
                // Find lvl node.
                var lvlNodes = abstractNumNode.Descendants().Where(n => n.Name.LocalName == "lvl");
                // No lvl, check if a numStyleLink is used.
                if (lvlNodes.Count() == 0) {
                    var linkedStyleNumId = GetLinkedStyleNumId(document, numIdValue);
                    if (linkedStyleNumId != -1) {
                        abstractNumNode = GetAbstractNum(document, linkedStyleNumId.ToString());
                        if (abstractNumNode != null) {
                            lvlNodes = abstractNumNode.Descendants().Where(n => n.Name.LocalName == "lvl");
                        }
                    }
                }
                XElement lvlNode = null;
                foreach (XElement node in lvlNodes) {
                    if (node.Attribute(w + "ilvl").Value.Equals(ilvlValue)) {
                        lvlNode = node;
                        break;
                    }
                    else if (ilvlValue == null) {
                        var numStyleNode = node.Descendants().FirstOrDefault(n => n.Name.LocalName == "pStyle");
                        if ((numStyleNode != null) && numStyleNode.GetAttribute(Document.w + "val").Equals(p.StyleId)) {
                            lvlNode = node;
                            break;
                        }
                    }
                }

                if (lvlNode != null) {
                    var numFmtNode = lvlNode.Descendants().FirstOrDefault(n => n.Name.LocalName == "numFmt");
                    if (numFmtNode != null)
                        return numFmtNode.Attribute(w + "val").Value;
                }
            }

            return null;
        }

        internal static string GetListItemStartValue(List list, int level) {
            var abstractNumElement = list.GetAbstractNum(list.NumId);
            if (abstractNumElement == null)
                return "1";

            //Find lvl node
            var lvlNodes = abstractNumElement.Descendants().Where(n => n.Name.LocalName == "lvl");
            var lvlNode = lvlNodes.FirstOrDefault(n => n.GetAttribute(Document.w + "ilvl").Equals(level.ToString()));
            // No ilvl, check if a numStyleLink is used.
            if (lvlNode == null) {
                var linkedStyleNumId = HelperFunctions.GetLinkedStyleNumId(list.Document, list.NumId.ToString());
                if (linkedStyleNumId != -1) {
                    abstractNumElement = list.GetAbstractNum(linkedStyleNumId);
                    if (abstractNumElement == null)
                        return "1";
                    lvlNodes = abstractNumElement.Descendants().Where(n => n.Name.LocalName == "lvl");
                    lvlNode = lvlNodes.FirstOrDefault(n => n.GetAttribute(Document.w + "ilvl").Equals(level.ToString()));
                }
                if (lvlNode == null)
                    return "1";
            }

            var startNode = lvlNode.Descendants().FirstOrDefault(n => n.Name.LocalName == "start");
            if (startNode == null)
                return "1";
            var returnValue = startNode.GetAttribute(Document.w + "val");


            var numNode = HelperFunctions.GetNumberingNumNode(list.Document, list.NumId.ToString());
            if (numNode != null) {
                var levelOverride = numNode.Elements(XName.Get("lvlOverride", Document.w.NamespaceName))?
                                           .SingleOrDefault(node => node.Attribute(Document.w + "ilvl").Value.Equals(level.ToString()));

                if (levelOverride != null) {
                    var startOverride = levelOverride.Element(XName.Get("startOverride", Document.w.NamespaceName));
                    if (startOverride != null) {
                        returnValue = startOverride.GetAttribute(XName.Get("val", Document.w.NamespaceName));
                    }
                }
            }

            return returnValue;
        }
        */

        //internal static string GetListItemNumberOptimized(this Paragraph par, Document doc) {
        //    if (par.IsListItem) {
        //        var list = doc.Lists.FirstOrDefault(l => l.NumId == par.GetNumId());
        //        if (list != null) {
        //            int listItemLevel = par.GetListItemLevel();
        //            if (listItemLevel != -1) {
        //                string listItemType = GetListItemType(par, doc);
        //                string listItemStartValue = HelperFunctions.GetListItemStartValue(list, listItemLevel);

        //                var textFormat = HelperFunctions.GetListItemTextFormat(list, listItemLevel, out var formatting);

        //                IEnumerable<Paragraph> source = list.Items.Where(p => p.GetListItemLevel() == listItemLevel);
        //                if (source.Count() > 0) {
        //                    int num = source.ToList().FindIndex((Paragraph p) => p.Text == par.Text);
        //                    if (num != -1) {
        //                        num = int.Parse(listItemStartValue) + num;
        //                        string text = "";
        //                        switch (listItemType) {
        //                            case "decimal":
        //                                text = num.ToString();
        //                                break;
        //                            case "lowerLetter":
        //                                text = ((char)((num - 1) % 26 + 97)).ToString();
        //                                break;
        //                            case "upperLetter":
        //                                text = ((char)((num - 1) % 26 + 65)).ToString();
        //                                break;
        //                            case "lowerRoman":
        //                                text = PdfParagraphPart.ConvertIntToRoman(num, isCapital: false);
        //                                break;
        //                            case "upperRoman":
        //                                text = PdfParagraphPart.ConvertIntToRoman(num, isCapital: true);
        //                                break;
        //                            default:
        //                                return null;
        //                        }

        //                        return textFormat;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    return null;
        //}

        internal static string ConvertIntToRoman(int number, bool isCapital) {
            if (number < 0 || number > 399) {
                throw new ArgumentOutOfRangeException("Value must be between 1 and 399");
            }

            if (number < 1) {
                return string.Empty;
            }

            if (number >= 100) {
                return (isCapital ? "C" : "c") + ConvertIntToRoman(number - 100, isCapital);
            }

            if (number >= 90) {
                return (isCapital ? "XC" : "xc") + ConvertIntToRoman(number - 90, isCapital);
            }

            if (number >= 50) {
                return (isCapital ? "L" : "l") + ConvertIntToRoman(number - 50, isCapital);
            }

            if (number >= 40) {
                return (isCapital ? "XL" : "xl") + ConvertIntToRoman(number - 40, isCapital);
            }

            if (number >= 10) {
                return (isCapital ? "X" : "x") + ConvertIntToRoman(number - 10, isCapital);
            }

            if (number >= 9) {
                return (isCapital ? "IX" : "ix") + ConvertIntToRoman(number - 9, isCapital);
            }

            if (number >= 5) {
                return (isCapital ? "V" : "v") + ConvertIntToRoman(number - 5, isCapital);
            }

            if (number >= 4) {
                return (isCapital ? "IV" : "iv") + ConvertIntToRoman(number - 4, isCapital);
            }

            if (number >= 1) {
                return (isCapital ? "I" : "i") + ConvertIntToRoman(number - 1, isCapital);
            }

            throw new ArgumentOutOfRangeException("Value must be between 1 and 399");
        }


        /// <summary>
        /// Кэширование документа (при необходимости) - для ускорения работы со списками
        /// </summary>
        /// <param name="doc"></param>
        internal static void CacheDoc(Document doc) {
            if (!m_cacheListParagraphs.TryGetValue(doc, out var parCache)) {
                parCache = new ConcurrentDictionary<Paragraph, (List list, int parIndex)>();
                var listCache = new ConcurrentDictionary<int, ConcurrentDictionary<int, Paragraph>>();

                if (doc.Lists?.Any() ?? false) {
                    foreach (var list in doc.Lists) {
                        var levelCache = new ConcurrentDictionary<int, Paragraph>();  //listCache[list.NumId] = list;
                        listCache[list.NumId] = levelCache;
                        var parNum = 0;
                        foreach (var par in list.Items) {
                            parCache[par] = (list, parNum++);
                            levelCache.TryAdd(par.IndentLevel ?? 0, par);
                        }
                    }
                }
                m_cacheListParagraphs.TryAdd(doc, parCache);
                m_cacheListLevels.TryAdd(doc, listCache);
            }
            //m_cacheLists.TryAdd(doc, listCache);
        }

        internal static string GetListItemNumberFast(this Paragraph par, Document doc, string textFormat = null) {
            string text = null;
            if (par.IsListItem && par.ListItemType == ListItemType.Numbered) {
                CacheDoc(doc);

                if (m_cacheListParagraphs.TryGetValue(doc, out var parCache)) {
                    if (parCache.TryGetValue(par, out var listInfo)) {
                        var levelConfig = listInfo.list.ListOptions.LevelsConfigs[par.IndentLevel.Value];
                        var num = listInfo.parIndex + (levelConfig.StartNumber ?? 1);
                        textFormat ??= levelConfig.NumberingLevelText;
                        
                        text = levelConfig.NumberingFormat switch {
                            NumberingFormat.decimalNormal => num.ToString(),
                            NumberingFormat.lowerLetter => ((char)((num - 1) % 26 + 97)).ToString(),
                            NumberingFormat.upperLetter => ((char)((num - 1) % 26 + 65)).ToString(),
                            NumberingFormat.lowerRoman => ConvertIntToRoman(num, isCapital: false),
                            NumberingFormat.upperRoman => ConvertIntToRoman(num, isCapital: true),
                            NumberingFormat.russianLower => ((char)((num - 1) % 33 + (int)'а')).ToString(),
                            NumberingFormat.russianUpper => ((char)((num - 1) % 33 + (int)'А')).ToString(),
                            _ => null
                        };
                        if (textFormat != null) {
                            int num2 = levelConfig.NumberingLevelText.LastIndexOf("%");
                            if (num2 >= 0) {
                                textFormat = textFormat.Remove(num2, 2).Insert(num2, text);
                            }
                            if (textFormat.Contains("%")) {
                                //ищем первый параграф уровнем выше
                                if (m_cacheListLevels.TryGetValue(doc, out var listCache) &&
                                    listCache.TryGetValue(listInfo.list.NumId, out var levelCache) &&
                                    levelCache.TryGetValue(par.IndentLevel.Value - 1, out var upperLevelPar)) {
                                    return upperLevelPar.GetListItemNumberFast(doc, textFormat);

                                }
                            }

                            return textFormat;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Получение текста параграфа с учётом номера элемента списка
        /// </summary>
        /// <param name="par"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        internal static string GetText(this Paragraph par, Document doc) {
            var parText = par.Text;

            if (par.IsListItem && par.ListItemType == ListItemType.Numbered) {
                //sw1.Start();
                //var numPrefix1 = par.GetListItemNumberFast(doc);
                //sw1.Stop();
                //sw2.Start();
                //var numPrefix2 = par.GetListItemNumber();
                //sw2.Stop();
                //if (!numPrefix1.Equals(numPrefix2)) {
                //var tt = 0;
                //}
                if (doc != null) {
                    SwFast.Start();
                    parText = $"{par.GetListItemNumberFast(doc)} {parText}";
                    SwFast.Stop();
                }
                else {
                    SwBase.Start();
                    parText = $"{par.GetListItemNumber()} {parText}";
                    SwBase.Stop();
                }
            }

            return parText;
        }
    }
}
