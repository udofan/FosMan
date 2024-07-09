using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    public class ErrorList : BaseObj {
        List<Error> m_errors = new();

        /// <summary>
        /// Список ошибок
        /// </summary>
        public List<Error> Items { get => m_errors; }

        /// <summary>
        /// Получение ошибки по индексу
        /// </summary>
        /// <param name="idx"></param>
        /// <returns></returns>
        public Error this[int idx] => m_errors[idx];

        /// <summary>
        /// Список пуст?
        /// </summary>
        public bool IsEmpty => m_errors.Count == 0;

        /// <summary>
        /// Счетчик
        /// </summary>
        public int Count => m_errors.Count;

        /// <summary>
        /// Добавить ошибку
        /// </summary>
        /// <param name="type"></param>
        /// <param name="message"></param>
        /// <param name="comment"></param>
        public void Add(EErrorType type, string comment = null) {
            var err = new Error() {
                Type = type,
                Message = type.GetDescription(),
                Comment = comment
            };
            m_errors.Add(err);
        }

        public void AddException(Exception ex) {
            Add(EErrorType.Exception, $"{ex.Message}\r\n{ex.StackTrace}");
        }

        /// <summary>
        /// Получить копию списка (объекты ошибок не клонируются, они по ссылке)
        /// </summary>
        /// <returns></returns>
        public List<Error> GetCopy() => m_errors?.ToList();

        public void AddRange(IEnumerable<Error> items) {
            if (items != null) {
                m_errors?.AddRange(items);
            }
        }

        /// <summary>
        /// Получение html-разметки с таблицей ошибок
        /// </summary>
        /// <returns></returns>
        public string GetHtmlReport() {
            var html = new StringBuilder();
            if (IsEmpty) {
                html.AddDiv("Список ошибок пуст.", "green");
            }
            else {
                html.AddDiv($"Список ошибок ({Count} шт.):", "red");

                var tdStyle = " style='border: 1px solid;'";
                var table = new StringBuilder(@$"<table {tdStyle}><tr style='font-weight: bold; background-color: lightgray'>");
                table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Тип</th><th {tdStyle}>Описание</th><th {tdStyle}>Комментарий</th>");
                table.Append("</tr>");
                var idx = 0;
                foreach (var item in Items) {
                    table.Append($"<tr><td {tdStyle}>{++idx}</td><td {tdStyle}>{item.Type}</td>" +
                                 $"<td {tdStyle}>{item.Type.GetDescription()}</td><td {tdStyle}>{item.Comment}</td></tr>");
                }
                table.Append("</table>");
                html.Append(table);
            }

            return html.ToString();
        }

        public void AddSimple(string text) {
            Add(EErrorType.Simple, text);
        }
    }
}
