using static FosMan.Enums;

namespace FosMan {
    public class Error {
        /// <summary>
        /// Код ошибки
        /// </summary>
        public EErrorType Type { set; get; }
        /// <summary>
        /// Сообщение
        /// </summary>
        public string Message { set; get; }
        /// <summary>
        /// Комментарий (html-разметка)
        /// </summary>
        public string Comment { set; get; }

        public override string ToString() => string.IsNullOrEmpty(Comment) ? $"{Message}" : $"{Message}: {Comment}";
    }
}
