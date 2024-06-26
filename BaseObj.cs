﻿using FastMember;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    /// <summary>
    /// Базовый объект с фичами аксессора к свойствам
    /// </summary>
    internal class BaseObj {
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
    }
}
