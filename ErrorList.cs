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
    }
}
