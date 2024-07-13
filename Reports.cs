using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static FosMan.Enums;
using static FosMan.App;

namespace FosMan {
    internal static class Reports {

        /// <summary>
        /// Сформировать отчет по описанию РПД
        /// </summary>
        public static string CreateRpdReportBase(List<Rpd> rpdList) {
            StringBuilder html = new($"<html><body><h2>Описание РПД</h2>");
            StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var sw = Stopwatch.StartNew();

            var idx = 0;
            foreach (var rpd in rpdList) {
                var anchor = $"rpd{idx}";
                idx++;

                rep.Append($"<div id='{anchor}' style='width: 100%;'><h3 style='background-color: lightsteelblue'>{rpd.DisciplineName ?? "?"}</h3>");
                rep.Append("<div style='padding-left: 30px'>");
                rep.AddFileLink($"Файл РПД:", rpd.SourceFileName);

                var table = rpd.GetPropertiesHtml();

                rep.Append(table);
                rep.Append("</div></div>");

                toc.AddTocElement(rpd.DisciplineName ?? "?", anchor, 0);
            }

            rep.Append("</div>");
            toc.Append("</ul></div>");
            html.AddDiv($"Дата: {DateTime.Now}");
            html.AddDiv($"РПД в списке: {rpdList.Count}");
            html.AddDiv($"Время работы: {sw.Elapsed}");
            html.Append("<p />");
            html.AddDiv($"<b>Список РПД:</b>");
            html.Append(toc).Append(rep).Append("</body></html>");

            return html.ToString();
        }

        /// <summary>
        /// Сформировать отчет по описанию РПД
        /// </summary>
        public static string CreateRpdReportFosMatching(List<Rpd> rpdList, List<Fos> fosList) {
            StringBuilder html = new($"<html><body><h2>Сопоставление РПД с ФОС</h2>");
            //StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var sw = Stopwatch.StartNew();

            var matchedPairs = new Dictionary<Rpd, Fos>();      //совпадающие пары (РПД-ФОС)
            var orphanedRpdList = new List<Rpd>();              //РПД без ФОСов
            var orphanedFosList = new List<Fos>();              //ФОСы без РПД

            //выявление сопоставлений
            var fosDic = fosList.ToDictionary(f => f.Key, f => f);
            foreach (var rpd in rpdList) {
                if (fosDic.TryGetValue(rpd.Key, out var fos)) {
                    matchedPairs[rpd] = fos;
                    fosDic.Remove(rpd.Key);
                }
                else {
                    orphanedRpdList.Add(rpd);
                }
            }
            orphanedFosList = fosDic.Values.ToList();

            rep.AddDiv($"<b>Успешно сопоставленные РПД с ФОС (шт.): {matchedPairs.Count}</b>", "green");

            //выдача списков сирот
            rep.Append("<p />");
            rep.AddDiv($"<b>Список РПД без ФОС ({orphanedRpdList.Count} шт.):</b>", "red");
            orphanedRpdList.Sort((rpd1, rpd2) => {
                var result = string.Compare($"{rpd1.DirectionCode}_{rpd1.Profile}", $"{rpd2.DirectionCode}_{rpd2.Profile}");
                if (result == 0) {
                    result = string.Compare(rpd1.Discipline?.DepartmentName, rpd2.Discipline?.DepartmentName);
                    if (result == 0) {
                        result = (rpd1.Discipline?.Number ?? 0) - (rpd2.Discipline?.Number ?? 0);
                    }
                }
                return result;
            });

            var tdStyle = " style='border: 1px solid;'";
            var table = new StringBuilder(@$"<table {tdStyle}><tr style='font-weight: bold; background-color: lightgray'>");
            table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Направление - Профиль</th><th {tdStyle}>Дисциплина</th><th {tdStyle}>Кафедра</th>");
            table.Append("</tr>");

            var idx = 0;
            foreach (var rpd in orphanedRpdList) {
                table.Append($"<tr {tdStyle}><td {tdStyle}>{++idx}</td><td {tdStyle}>{rpd.Discipline?.Curriculum?.DirectionCode ?? rpd.DirectionCode} " +
                             $"{rpd.Discipline?.Curriculum?.DirectionName ?? rpd.DirectionName} - {rpd.Discipline?.Curriculum?.Profile ?? rpd.Profile}</td>" +
                             $"<td {tdStyle}>{rpd.CurriculumDisciplineName ?? rpd.DisciplineName}</td>" +
                             $"<td {tdStyle}>{rpd.Discipline?.DepartmentName ?? rpd.Department}</td>" +
                             $"</tr>");
            }
            table.Append("</table>");
            rep.Append(table);

            //список ФОС без РПД
            rep.Append("<p />");
            rep.AddDiv($"<b>Список ФОС без РПД ({orphanedFosList.Count} шт.):</b>", "red");
            orphanedFosList.Sort((fos1, fos2) => {
                var result = string.Compare($"{fos1.DirectionCode}_{fos1.Profile}", $"{fos2.DirectionCode}_{fos2.Profile}");
                if (result == 0) {
                    result = string.Compare(fos1.Discipline?.DepartmentName, fos2.Discipline?.DepartmentName);
                    if (result == 0) {
                        result = (fos1.Discipline?.Number ?? 0) - (fos2.Discipline?.Number ?? 0);
                    }
                }
                return result;
            });

            var table2 = new StringBuilder(@$"<table {tdStyle}><tr style='font-weight: bold; background-color: lightgray'>");
            table2.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Направление - Профиль</th><th {tdStyle}>Дисциплина</th><th {tdStyle}>Кафедра</th>");
            table2.Append("</tr>");

            idx = 0;
            foreach (var fos in orphanedFosList) {
                table2.Append($"<tr {tdStyle}><td {tdStyle}>{++idx}</td><td {tdStyle}>{fos.Discipline?.Curriculum?.DirectionCode ?? fos.DirectionCode} " +
                              $"{fos.Discipline?.Curriculum?.DirectionName ?? fos.DirectionName} - {fos.Discipline?.Curriculum?.Profile ?? fos.Profile}</td>" +
                              $"<td {tdStyle}>{fos.Discipline?.Name ?? fos.DisciplineName}</td>" +
                              $"<td {tdStyle}>{fos.Discipline?.DepartmentName ?? fos.Department}</td>" +
                              $"</tr>");
            }
            table2.Append("</table>");
            rep.Append(table2);

            rep.Append("</div>");
            //toc.Append("</ul></div>");
            html.AddDiv($"Дата: {DateTime.Now}");
            html.AddDiv($"РПД (шт.): {rpdList.Count}");
            html.AddDiv($"ФОС (шт.): {fosList.Count}");
            html.AddDiv($"Время работы: {sw.Elapsed}");
            html.Append("<p />");
            //html.AddDiv($"<b>Список РПД:</b>");
            //html.Append(toc).Append(rep).Append("</body></html>");
            html.Append(rep).Append("</body></html>");

            return html.ToString();
        }

        /// <summary>
        /// Сформировать html-разметку с таблицей дисциплин из УП для блока
        /// </summary>
        /// <param name="curriculum"></param>
        /// <param name="disciplineType"></param>
        /// <param name="blockNum"></param>
        /// <param name="rpdList"></param>
        /// <returns></returns>
        static string CreateHtmlForDisciplinesForCurriculumBlock(Curriculum curriculum, EDisciplineType disciplineType, int blockNum, List<Rpd> rpdList, List<Fos> fosList) {
            var disciplines = curriculum.Disciplines.Values
                                .Where(d => d.Type == disciplineType && d.BlockNum == blockNum)
                                .OrderBy(x => x.Number)
                                .ToList();

            var html = new StringBuilder($"<h3>Дисциплины типа [{disciplineType.GetDescription()}] ({disciplines.Count} шт.):</h3>");
            var table = CreateHtmlTableForDisciplines(disciplines, rpdList, fosList);
            html.Append(table);

            return html.ToString();
        }

        /// <summary>
        /// Сформировать html-разметку с таблицей для списка дисциплин
        /// </summary>
        /// <param name="curriculum"></param>
        /// <param name="disciplineType"></param>
        /// <param name="blockNum"></param>
        /// <param name="rpdList"></param>
        /// <returns></returns>
        static string CreateHtmlTableForDisciplines(IEnumerable<CurriculumDiscipline> disciplines, List<Rpd> rpdList, List<Fos> fosList) {
            var html = new StringBuilder();

            var tdStyle = " style='border: 1px solid;'";
            var table = new StringBuilder(@"<table style='border: 1px solid'><tr style='font-weight: bold; background-color: lightgray'>");
            table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Индекс</th><th {tdStyle}>Название</th><th {tdStyle}>Кафедра</th>" +
                         $"<th {tdStyle}>Наличие РПД</th><th {tdStyle}>Наличие ФОС</th>");
            table.Append("</tr>");

            var tdStyleCheck = "style='border: 1px solid; background-color: lightgreen; text-align: center'";
            var tdStyleTimes = "style='border: 1px solid; background-color: lightpink; text-align: center'";
            var spanCheck = "<span style='color: green'>&check;</span>";
            var spanTimes = "<span style='color: red'>&times;</span>";

            var idx = 0;
            foreach (var disc in disciplines) {
                var rpd = App.FindRpd(disc, rpdList);
                var fos = App.FindFos(disc, fosList);
                //var discNameStyle = "style='background-color: lightgreen;'";
                var rpdStyle = rpd != null ? tdStyleCheck : tdStyleTimes;
                var fosStyle = fos != null ? tdStyleCheck : tdStyleTimes;
                var rpdMatch = rpd != null ? spanCheck : spanTimes;
                var fosMatch = fos != null ? spanCheck : spanTimes;

                table.Append($"<tr {tdStyle}><td {tdStyle}>{++idx}</td><td {tdStyle}>{disc.Index}</td><td {tdStyle}>{disc.Name}</td>" +
                             $"<td {tdStyle}>{disc.DepartmentName}</td><td {rpdStyle}>{rpdMatch}</td><td {fosStyle}>{fosMatch}</td></tr>");
            }

            table.Append("</table>");
            html.Append(table);

            return html.ToString();
        }

        /// <summary>
        /// Проверка РПД и ФОС по УП с отчетом
        /// </summary>
        /// <param name="rpdList"></param>
        /// <returns></returns>
        public static string CreateReportRpdAndFosCheckByCurricula(List<Rpd> rpdList, List<Fos> fosList) {
            StringBuilder html = new($"<html><body><h2>Проверка РПД и ФОС по УП</h2>");
            //StringBuilder toc = new("<div><ul>");
            StringBuilder rep = new("<div>");
            var sw = Stopwatch.StartNew();

            //обязательные дисциплины
            var curricula = App.FindCurricula(rpdList.FirstOrDefault());
            var curriculum = curricula?.Values.FirstOrDefault(c => c.FormOfStudy == EFormOfStudy.FullTime);
            if (curriculum == null) {
                rep.AddError("УП не найдены. Их загрузка осуществляется на вкладке <b>\"Учебные планы\"</b>.");
            }
            else {
                rep.AddDiv($"Направление: {curriculum.DirectionCode} {curriculum.DirectionName}");
                rep.AddDiv($"Профиль: {curriculum.Profile}");

                rep.Append(CreateHtmlForDisciplinesForCurriculumBlock(curriculum, EDisciplineType.Required, 1, rpdList, fosList));
                rep.Append(CreateHtmlForDisciplinesForCurriculumBlock(curriculum, EDisciplineType.Variable, 1, rpdList, fosList));
                rep.Append(CreateHtmlForDisciplinesForCurriculumBlock(curriculum, EDisciplineType.ByChoice, 1, rpdList, fosList));
                rep.Append(CreateHtmlForDisciplinesForCurriculumBlock(curriculum, EDisciplineType.Optional, -1, rpdList, fosList));

                var depCodes = curriculum.Disciplines.Values.Select(d => d.DepartmentCode).ToHashSet();

                foreach (var depCode in depCodes) {
                    var depDisciplines = curriculum.Disciplines.Values.Where(d => d.DepartmentCode.Equals(depCode)).OrderBy(d => d.Number).ToList();
                    rep.Append("<p />");
                    rep.AddDiv($"<h3>Дисциплины кафедры <b>{depDisciplines.FirstOrDefault().DepartmentName}</b> [{depCode}] ({depDisciplines.Count} шт.):</h3>");
                    var discTable = CreateHtmlTableForDisciplines(depDisciplines, rpdList, fosList);
                    rep.Append(discTable);
                }
            }
            rep.Append("</div>");
            //toc.Append("</ul></div>");
            html.AddDiv($"Дата: {DateTime.Now}");
            html.AddDiv($"РПД в списке: {rpdList.Count}");
            html.AddDiv($"ФОС в списке: {fosList.Count}");
            html.AddDiv($"Время работы: {sw.Elapsed}");
            html.Append("<p />");
            html.Append(rep).Append("</body></html>");

            return html.ToString();
        }
    }
}
