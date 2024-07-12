using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static FosMan.Enums;

namespace FosMan {
    internal class FileFixerItem {
        public string DirectoryName { get; set; }
        public string FileName { get; set; }
        public string FullFileName { get; set; }
        public EFileType Type { get; set; }
        public Rpd Rpd { get; set; }
        public Fos Fos { get; set; }

        public FileFixerItem(string file, EFileType fileType) {
            FullFileName = file;
            DirectoryName = Path.GetDirectoryName(file);
            FileName = Path.GetFileName(file);

            if (fileType == EFileType.Auto) {
                if (FileName.Contains("РПД")) {
                    fileType = EFileType.Rpd;
                }
                else if (FileName.Contains("ФОС")) {
                    fileType = EFileType.Fos;
                }
                else {
                    fileType = EFileType.Unknown;
                }
            }

            Type = fileType;
        }

        public FileFixerItem(Rpd rpd) {
            Rpd = rpd;
            Type = EFileType.Rpd;
            FullFileName = rpd.SourceFileName;
            DirectoryName = Path.GetDirectoryName(FullFileName);
            FileName = Path.GetFileName(FullFileName);
        }
        public FileFixerItem(Fos fos) {
            Fos = fos;
            Type = EFileType.Fos;
            FullFileName = fos.SourceFileName;
            DirectoryName = Path.GetDirectoryName(FullFileName);
            FileName = Path.GetFileName(FullFileName);
        }

        /// <summary>
        /// Поиск вхождений в файле
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        internal int FindAllEntries(string pattern, out string report) {
            var html = new StringBuilder();
            var matchCount = 0;
            var sw = Stopwatch.StartNew();
            var table = new StringBuilder();
            try {
                var tdStyle = " style='border: 1px solid;'";
                table.Append($"<table {tdStyle}><tr style='font-weight: bold; background-color: lightgray'>");
                table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Вхождение</th></tr>");

                var replaceValue = "<span style='background-color: blue; color: white'>$+</span>";

                //var re = new Regex(pattern, RegexOptions.IgnoreCase);
                using (var docx = DocX.Load(this.FullFileName)) {
                    foreach (var par in docx.Paragraphs) {
                        var replaceTextOptions = new FunctionReplaceTextOptions() {
                            FindPattern = pattern,
                            ContainerLocation = ReplaceTextContainer.All,
                            RegExOptions = RegexOptions.IgnoreCase,
                            RegexMatchHandler = m => {
                                var text = par.PreviousParagraph?.Text ?? "";
                                if (!string.IsNullOrEmpty(text)) text += " ";
                                var parText = Regex.Replace(par.Text, 
                                                            pattern, 
                                                            replaceValue, 
                                                            RegexOptions.IgnoreCase);
                                text += parText;
                                if (par.NextParagraph != null) {
                                    text += $" {par.NextParagraph.Text}";
                                }

                                table.Append($"<tr {tdStyle}><td {tdStyle}>{++matchCount}</td><td {tdStyle}>{text}</td></tr>");
                                return m;
                            }
                        };
                        par.ReplaceText(replaceTextOptions);
                    }
                }
                table.Append("</table>");
            }
            catch (Exception ex) {
                html.AddError($"{ex.Message}<br />{ex.StackTrace}");
            }

            if (matchCount == 0) {
                html.AddDiv($"Вхождений не найдено.");
            }
            else {
                html.AddDiv($"Найденные вхождения ({matchCount} шт.):");
                html.Append(table);
            }

            html.AddDiv($"Время работы: {sw.Elapsed}");
            report = html.ToString();
            return matchCount;
        }
    }
}
