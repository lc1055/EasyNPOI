using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI;
using NPOI.XWPF.UserModel;

namespace EasyNPOI.Services
{
    public class WordExportService : IWordExportService
    {
        /// <summary>
        /// 根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tmplPath"></param>
        /// <param name="savePath"></param>
        /// <param name="wordData"></param>
        public void ExportByTemplate<T>(string tmplPath, string savePath, T wordData)  where T : class , new()
        {
            if(string.IsNullOrEmpty(tmplPath) || string.IsNullOrEmpty(savePath))
            {
                throw new ArgumentNullException("路径为空");
            }

            XWPFDocument word = NPOIHelper.GetXWPFDocument(tmplPath);
            ReplacePlaceholders(word, wordData);
            NPOIHelper.SaveXWPFDocument(savePath, word);
        }


        /// <summary>
        /// 替换模板中的占位符
        /// </summary>
        /// <param name="word"></param>
        private void ReplacePlaceholders<T>(XWPFDocument word, T wordData) where T : class, new()
        {
            if (word == null)
            {
                throw new ArgumentNullException("XWPFDocument 对象为空");
            }

            var basicReplacements = PlaceholderHelper.GetBasicReplacements(wordData);
            var gridReplacements = PlaceholderHelper.GetGridReplacements(wordData);

            NPOIHelper.ReplaceInWord(word, basicReplacements, gridReplacements);
        }


    }
}
