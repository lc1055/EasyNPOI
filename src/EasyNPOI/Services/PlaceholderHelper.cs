using EasyNPOI.Attributes;
using EasyNPOI.Models.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Services
{
    public class PlaceholderHelper
    {
        /// <summary>
        /// 将单个属性生成替换对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="prop"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        private static ReplacementBasic GetBasicReplacementByProperty<T>(PropertyInfo prop, T data)
        {
            var placeholder = prop.Name;
            var attr = prop.GetCustomAttribute<Attributes.PlaceholderAttribute>();
            if (attr != null)
            {
                if (!string.IsNullOrEmpty(attr.PlaceholderName))
                {
                    placeholder = attr.PlaceholderName;
                }
            }
            placeholder = "{" + placeholder + "}";

            //图片
            if (prop.PropertyType == typeof(Models.Word.Picture) || typeof(IEnumerable<Models.Word.Picture>).IsAssignableFrom(prop.PropertyType))
            {
                List<Picture> pictures = new List<Picture>();
                if (prop.PropertyType == typeof(Picture))
                {
                    var picture = (Picture)prop.GetValue(data);
                    //将单张图片对象也统一成列表
                    pictures = new List<Picture>() { picture };
                }

                if (typeof(List<Picture>).IsAssignableFrom(prop.PropertyType))
                {
                    pictures = (List<Picture>)prop.GetValue(data);              
                }
                return new ReplacementBasic { Placeholder = placeholder, Pictures = pictures, Type = Enums.PlaceholderTypeEnum.Picture };
            }
            //表格
            else if (prop.PropertyType.IsDefined(typeof(PlaceholderGridAttribute)))
            {
                return null;
            }
            //文本
            else
            {
                var replaceValue = prop.GetValue(data)?.ToString();
                return new Models.Word.ReplacementBasic { Placeholder = placeholder, Text = replaceValue, Type = Enums.PlaceholderTypeEnum.Text };
            }
        }


        /// <summary>
        /// 生成基础替换对象(文本/图片)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="wordData"></param>
        /// <returns></returns>
        public static List<Models.Word.ReplacementBasic> GetBasicReplacements<T>(T wordData) where T : class, new()
        {
            var replacements = new List<Models.Word.ReplacementBasic>();
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();
            foreach (PropertyInfo prop in props)
            {
                var replace = GetBasicReplacementByProperty(prop, wordData);
                if (replace != null)
                {
                    replacements.Add(replace);
                }
            }
            return replacements;
        }


        /// <summary>
        /// 生成表格替换对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="wordData"></param>
        /// <returns></returns>
        public static List<Models.Word.ReplacementGrid> GetGridReplacements<T>(T wordData) where T : class, new()
        {
            var replacements = new List<Models.Word.ReplacementGrid>();
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();

            foreach (PropertyInfo prop in props)
            {
                if (prop.IsDefined(typeof(PlaceholderGridAttribute)))
                {
                    var placeholder = prop.Name;
                    var attr = prop.GetCustomAttribute<Attributes.PlaceholderGridAttribute>();
                    if (attr != null)
                    {
                        if (!string.IsNullOrEmpty(attr.PlaceholderName))
                        {
                            placeholder = attr.PlaceholderName;
                        }
                    }
                    placeholder = "{{" + placeholder + "}}";

                    if (typeof(List<>).IsAssignableFrom(prop.PropertyType.GetGenericTypeDefinition()))
                    {
                        var objectList = prop.GetValue(wordData);
                        IEnumerable<object> list = null;
                        try
                        {
                            list = objectList as IEnumerable<object>;
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                        if (list == null || list.Count() == 0)
                        {
                            continue;
                        }

                        List<ReplacementRow> rows = new List<ReplacementRow>();
                        foreach (var o in list)
                        {
                            var row = new ReplacementRow();
                            PropertyInfo[] _props = o.GetType().GetProperties();
                            foreach (PropertyInfo pi in _props)
                            {
                                var replace = GetBasicReplacementByProperty(pi, o);
                                row.Cells.Add(replace);
                            }
                            rows.Add(row);
                        }
                        replacements.Add(new ReplacementGrid { Placeholder = placeholder, Rows = rows });

                    }
                }
            }

            return replacements;
        }


    }
}
