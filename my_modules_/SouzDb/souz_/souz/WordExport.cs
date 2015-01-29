using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Microsoft.Office.Interop.Word;

namespace souz.DataExport
    {
    public static class WordExport
        {
        struct keyWordEntry
            {
            public string keyword;
            public int position;
            public string spacesAfter;

            public keyWordEntry(string kword, int pos, string spaces)
                {
                keyword = kword;
                position = pos;
                spacesAfter = spaces;
                }
            }


        public static void Export(System.Data.DataTable dt, string templPath, string destPath)
            {
            //Объекты для работы с вордаом
            //заглушка для опциональных аргументов
            object oMissing = System.Reflection.Missing.Value;
            //разделитель страниц http://msdn.microsoft.com/en-us/library/bb213704%28office.12%29.aspx
            object pageBreak = WdBreakType.wdPageBreak;
            //не сохранять изменения
            object noSave = WdSaveOptions.wdDoNotSaveChanges;

            //путь к шаблону
            object template = templPath;

            //куда сохранять полученный документ
            object destination = destPath;

            //запускаем Word
            _Application word = new Microsoft.Office.Interop.Word.Application();


            //Можем сделать его видимым и смотреть как скачут слова, абзацы и страницы
            word.Visible = true;

            //Создаем временный документ, в котором будем заменять ключевые слова на наши
            _Document sdoc = word.Documents.Add(ref template, ref oMissing, ref oMissing, ref oMissing);

            //загружаем ключевые слова
            string[] keyWords = { "FIO", "SUMMA", "DATEOPL", "RASHODY", "ITOGO" };

            //Ищем позиции ключевых слов в документе и добавляем в список
            List<keyWordEntry> keyWordEntries = new List<keyWordEntry>();
            for (int i = 0; i < sdoc.Words.Count; i++)
                {
                foreach (string keyWord in keyWords)
                    {
                    if (sdoc.Words[i + 1].Text.Trim() == keyWord) //не забываем, что ворд считает с единицы, а не нуля
                        {
                        keyWordEntries.Add(new keyWordEntry(keyWord, i + 1, sdoc.Words[i + 1].Text.Remove(0, keyWord.Length)));
                        };
                    };
                };


            //Создаем документ назначения, на основе шаблона, чтобы сохранилась разметка страницы, стили, колонтитулы и т.п.
            _Document ddoc = word.Documents.Add(ref template, ref oMissing, ref oMissing, ref oMissing);
            //Удаляем из него все тексты картинки и т.п.
            ddoc.Range(ref oMissing, ref oMissing).Delete(ref oMissing, ref oMissing);

            int rowCount = dt.Rows.Count;

            //Размечаем документ по количеству записей
            for (int i = 0; i < rowCount; i++)
                {
                ddoc.Range(ref oMissing, ref oMissing).InsertParagraphAfter();
                };

            //заполняем документ с конца
            for (int i = rowCount; i > 0; i--)
                {
                if (i < rowCount)
                    {
                    ddoc.Paragraphs[i].Range.InsertParagraphAfter();
                    ddoc.Paragraphs[i + 1].Range.InsertBreak(ref pageBreak);
                    };
                //подставляем слова во временный документ
                foreach (keyWordEntry ke in keyWordEntries)
                    {
                    string replaceWith = "";
                    switch (ke.keyword)
                        {
                        case "FNAME":
                            replaceWith = dt.Rows[i - 1]["firstname"].ToString() + ke.spacesAfter;
                            break;
                        case "LNAME":
                            replaceWith = dt.Rows[i - 1]["lastname"].ToString() + ke.spacesAfter;
                            break;
                        case "DEBT":
                            replaceWith = dt.Rows[i - 1]["debt"].ToString() + ke.spacesAfter;
                            break;
                        case "MR":
                            if (dt.Rows[i - 1]["gender"] == "M")
                                {
                                replaceWith = "Mr" + ke.spacesAfter;
                                }
                            else
                                {
                                replaceWith = "Mrs" + ke.spacesAfter;
                                };
                            break;
                        default:
                            replaceWith = ke.keyword + ke.spacesAfter;
                            break;
                        };
                    sdoc.Words[ke.position].Text = replaceWith;
                    };
                sdoc.Range(ref oMissing, ref oMissing).Copy();
                ddoc.Paragraphs[i].Range.Paste();
                }

            //закрываем временный документ без сохранения
            sdoc.Close(ref noSave, ref oMissing, ref oMissing);
            //сахраняем полученный документ
            ddoc.SaveAs(ref destination, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //закрываем полученный документ
            ddoc.Close(ref oMissing, ref oMissing, ref oMissing);
            //завершаем наш процесс ворда
            word.Quit(ref oMissing, ref oMissing, ref oMissing);
            }



        } // конец класса WordExport

       
    
}
