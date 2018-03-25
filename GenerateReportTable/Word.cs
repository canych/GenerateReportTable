using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace GenerateReportTable
{
    class WordDoc
    {
        /// <summary>
        /// Приложение Word
        /// </summary>
        private Word.Application _wordApp;

        /// <summary>
        /// Конструктор
        /// </summary>
        public WordDoc()
        {
            // Создание нового приложения
            _wordApp = new Word.Application();
        }

        /// <summary>
        /// Закрытие приложения
        /// </summary>
        public void Close()
        {
            // Запрос на сохранение документа
            Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges;
            // Формат сохранения
            Object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            // Необязательный параметр. При true документ направляется следующему получателю,
            // если документ является attached документом
            Object routeDocument = Type.Missing;

            // Выход
            _wordApp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
            _wordApp = null;
        }
    }
}
