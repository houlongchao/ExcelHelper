using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelHelper
{
    /// <summary>
    /// Excel对象信息
    /// </summary>
    public class ExcelObjectInfo
    {
        /// <summary>
        /// 对象类型
        /// </summary>
        public Type ObjectType { get; }

        #region 导入

        /// <summary>
        /// 导入唯一限制
        /// </summary>
        public IEnumerable<ImportUniqueAttribute> ImportUniques { get; set; }

        #endregion

        /// <summary>
        /// Excel 对象信息
        /// </summary>
        /// <param name="objectType"></param>
        public ExcelObjectInfo(Type objectType)
        {
            ObjectType = objectType;
            ImportUniques = ObjectType.GetCustomAttributes<ImportUniqueAttribute>();
        }

        /// <summary>
        /// 唯一判断字典
        /// </summary>
        private Dictionary<string, HashSet<string>> uniqueDict = new Dictionary<string, HashSet<string>>();

        /// <summary>
        /// 检查导入唯一性限制
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="t"></param>
        /// <exception cref="ImportException"></exception>
        public void CheckImportUnique<T>(T t) where T : new()
        {
            if (ImportUniques == null)
            {
                return;
            }

            foreach (var importUnique in ImportUniques)
            {
                if (importUnique.UniquePropertites.Length <= 0)
                {
                    continue;
                }

                var key = string.Join(",", importUnique.UniquePropertites);

                if (!uniqueDict.ContainsKey(key))
                {
                    uniqueDict[key] = new HashSet<string>();
                }

                string value = "";
                foreach (var uniqueProperty in importUnique.UniquePropertites)
                {
                    var data = typeof(T).GetProperty(uniqueProperty)?.GetValue(t)?.ToString();
                    value += data + ";";
                }

                if (uniqueDict[key].Contains(value))
                {
                    throw new ImportException($"数据导入唯一性限制【{key}】");
                }

                uniqueDict[key].Add(value);
            }
        }
    }
}
