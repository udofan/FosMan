using FastMember;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using Xceed.Pdf.Atoms;

namespace FosMan {
    /// <summary>
    /// Базовый объект с фичами аксессора к свойствам
    /// </summary>
    public class BaseObj {
        static ConcurrentDictionary<Type, TypeAccessor> m_typeAccessors = []; // TypeAccessor.Create(typeof(CurriculumDiscipline));
        static ConcurrentDictionary<Type, TypeAccessor> m_extraTypeAccessors = [];

        TypeAccessor TypeAccessor { get => m_typeAccessors.GetOrAdd(this.GetType(), TypeAccessor.Create(this.GetType())); }

        /// <summary>
        /// Получить значение свойства по имени
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        public object GetProperty(string propName) {
            //m_typeAccessor ??= TypeAccessor.Create(this.GetType());

            object value = null;
            //object instance = this;
            try {
                //проверка на квалифицированное имя (с точкой)
                var dotPos = propName.IndexOf('.');
                if (dotPos > 0) {
                    var basePropName = propName[..dotPos];
                    var instance = TypeAccessor[this, basePropName] as BaseObj;
                    value = instance?.GetProperty(propName[(dotPos + 1)..]);
                }
                else {
                    value = TypeAccessor[this, propName];
                }
            }
            catch (Exception ex) {
            }

            return value;
        }

        public MemberSet GetPropertySet() => TypeAccessor?.GetMembers();

        /// <summary>
        /// Установить значение указанному свойству
        /// </summary>
        /// <param name="targetProperty">целевое свойство</param>
        /// <param name="targetType">тип целевого свойства</param>
        /// <param name="propValue">значение (будет приведено к типу targetType)</param>
        /// <param name="index">индекс списка, если целевое свойство List(object)</param>
        /// <param name="subProperty">свойство object'а из списка</param>
        internal void SetProperty(string targetProperty, Type targetType, string propValue, int index = -1, string subProperty = null) {
            object value = null;

            propValue = propValue?.Trim();
            //value = propValue;

            if (targetType == typeof(int)) {
                if (int.TryParse(propValue, out var intValue)) {
                    value = intValue;
                }
            }
            else if (targetType == typeof(string)) {
                value = propValue;
            }
            else {
                value = propValue;
            }

            if (index >= 0 && !string.IsNullOrEmpty(subProperty)) {
                var arrayObj = TypeAccessor[this, targetProperty] as object[];
                if (arrayObj != null) {
                    var extraType = arrayObj.FirstOrDefault()?.GetType();
                    if (extraType != null) {
                        var extraTypeAccessor = m_extraTypeAccessors.GetOrAdd(extraType, TypeAccessor.Create(extraType));
                        extraTypeAccessor[arrayObj[index], subProperty] = value;
                    }
                }
            }
            else {
                TypeAccessor[this, targetProperty] = value;
            }
        }

        internal string GetPropertyHtml(string propName) {
            var value = "";

            var propValue = GetProperty(propName);
            if (propValue is BaseObj propBaseObj) {
                value = propBaseObj.GetPropertiesHtml();
            }
            else {
                if (propValue is IDictionary dict) {
                    foreach (var key in dict.Keys) {
                        //if (value.Length > 0) value += "<br />";
                        
                        var item = dict[key];
                        if (item is BaseObj itemBaseObj) {
                            value += $"<details open><summary>{key}</summary>";
                            value += itemBaseObj.GetPropertiesHtml();
                            value += "</details>";
                        }
                        else {
                            if (value.Length > 0) value += "<br />";
                            value += $"{key}: {item ?? "null"}";
                        }
                        
                    }
                }
                else if (!(propValue is string) && (propValue is IEnumerable list)) { // list || propValue is ISet set) {
                    var idx = 0;
                    foreach (var item in list) {
                        //if (value.Length > 0) value += "<br />";
                        if (item is BaseObj itemBaseObj) {
                            value += $"<details open><summary>{idx++}</summary>";
                            value += itemBaseObj.GetPropertiesHtml();
                            value += "</details>";
                        }
                        else {
                            if (value.Length > 0) value += "<br />";
                            value += $"{idx++}: {item ?? "null"}";
                        }
                    }
                }
                else {
                    value = propValue?.ToString() ?? "null";
                }
            }

            return value;
        }

        /// <summary>
        /// Формирование html-разметки со значениями свойств
        /// </summary>
        /// <returns></returns>
        internal string GetPropertiesHtml() {
            var tdStyle = " style='border: 1px solid;'";
            var table = new StringBuilder(@$"<table {tdStyle}><tr style='font-weight: bold; background-color: lightgray'>");
            table.Append($"<th {tdStyle}>№ п/п</th><th {tdStyle}>Свойство</th><th {tdStyle}>Значение</th>");
            table.Append("</tr>");

            var props = GetPropertySet(); // rpd.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
            var propIdx = 0;
            foreach (var prop in props) {
                var propValue = GetPropertyHtml(prop.Name);

                table.Append($"<tr><td {tdStyle}>{++propIdx}</td><td {tdStyle}><b>{prop.Name}</b></td><td {tdStyle}>{propValue}</td></tr>");
                //    propIdx++;
            }
            table.Append("</table>");

            return table.ToString();
        }
    }
}
