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
        /// Документ
        /// </summary>
        private Word.Document _wordDocument;
        /// <summary>
        /// Коллекция абзацев
        /// </summary>
        private Word.Paragraphs _wordParagraphs;
        /// <summary>
        /// Абзац
        /// </summary>
        private Word.Paragraph _wordParagraph;

        /// <summary>
        /// Конструктор
        /// </summary>
        public WordDoc()
        {
            // Создание нового приложения
            _wordApp = new Word.Application();

            // Создание нового документа
            _wordDocument = _wordApp.Documents.Add();
        }

        public void CreateTableName()
        {
            //Получаем ссылки на параграфы документа
            _wordParagraphs = _wordDocument.Paragraphs;
            // Будем работать с первым параграфом
            _wordParagraph = _wordParagraphs[1];
            // Выводим текст в первый параграф
            _wordParagraph.Range.Text = "Таблица 1.X – Рецензия-рейтинг на проведение занятия со студентами при прохождении научно-педагогической практики";
            // Меняем характеристики текста и параграфа
            _wordParagraph.Range.Font.Color = Word.WdColor.wdColorBlack;
            _wordParagraph.Range.Font.Size = 12;
            _wordParagraph.Range.Font.Name = "TimesNewRoman";
            _wordParagraph.Range.Font.Italic = 0;
            _wordParagraph.Range.Font.Bold = 0;
            // Абзацный отступ
            _wordParagraph.FirstLineIndent = 0;
            // Выравнивание
            _wordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        }

        public void CreateTable(List<Report> list)
        {
            // Добавляем в документ несколько параграфов
            _wordDocument.Paragraphs.Add(Missing.Value);
        }

        /// <summary>
        /// Сохранение документа
        /// </summary>
        /// <param name="path">Путь для сохранения</param>
        public void Save(string path)
        {
            _wordDocument.SaveAs($"{path}");
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
            _wordApp.Quit(saveChanges, originalFormat, routeDocument);
            _wordApp = null;
        }
    }
}
