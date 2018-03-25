using System;
using System.Collections.Generic;
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
            // Получаем ссылки на параграфы документа
            _wordParagraphs = _wordDocument.Paragraphs;
            // Будем работать с первым параграфом
            _wordParagraph = _wordParagraphs[1];
            // Выводим текст в первый параграф
            _wordParagraph.Range.Text = "Таблица 1.X – Рецензия-рейтинг на проведение занятия со студентами при прохождении научно-педагогической практики";
            // Меняем характеристики текста и параграфа
            _wordParagraph.Range.Font.Color = Word.WdColor.wdColorBlack;
            _wordParagraph.Range.Font.Size = 12;
            _wordParagraph.Range.Font.Name = "Times New Roman";
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

            // Получаем ссылки на параграфы документа
            _wordParagraphs = _wordDocument.Paragraphs;
            // Будем работать со вторым параграфом
            _wordParagraph = _wordParagraphs[2];

            // Новая таблица
            Word.Table _wordTable = _wordDocument.Tables.Add(_wordParagraph.Range, list.Count, 7,
                Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitWindow);

            #region Ширина столбцов
            _wordTable.Columns[1].SetWidth(ColumnWidth: 28f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            _wordTable.Columns[2].SetWidth(ColumnWidth: 192f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            _wordTable.Columns[3].SetWidth(ColumnWidth: 28f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            _wordTable.Columns[4].SetWidth(ColumnWidth: 28f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            _wordTable.Columns[5].SetWidth(ColumnWidth: 28f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            _wordTable.Columns[6].SetWidth(ColumnWidth: 28f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            _wordTable.Columns[7].SetWidth(ColumnWidth: 135f, RulerStyle: Word.WdRulerStyle.wdAdjustNone);
            #endregion

            #region Объединение ячеек
            // № п/п
            object begCell = _wordTable.Cell(1, 1).Range.Start;
            object endCell = _wordTable.Cell(2, 1).Range.End;
            Word.Range wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.Cells.Merge();

            // Критерии оценки
            begCell = _wordTable.Cell(1, 2).Range.Start;
            endCell = _wordTable.Cell(2, 2).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.Cells.Merge();

            // Шкала оценок
            begCell = _wordTable.Cell(1, 3).Range.Start;
            endCell = _wordTable.Cell(1, 6).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.Cells.Merge();

            // Краткое обоснование оценки
            begCell = _wordTable.Cell(1, 4).Range.Start;
            endCell = _wordTable.Cell(2, 7).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.Cells.Merge();
            #endregion

            #region Заполнение таблицы
            // Шапка
            Word.Range _wordCellRange = _wordTable.Cell(1, 1).Range;
            _wordCellRange.Text = "№ п/п";
            _wordCellRange = _wordTable.Cell(1, 2).Range;
            _wordCellRange.Text = "Критерии оценки";
            _wordCellRange = _wordTable.Cell(1, 3).Range;
            _wordCellRange.Text = "Шкала оценок";
            _wordCellRange = _wordTable.Cell(1, 4).Range;
            _wordCellRange.Text = "Краткое обоснование\nоценки";

            // Виды оценок
            for (int i = 3; i < 7; i++)
            {
                _wordCellRange = _wordTable.Cell(2, i).Range;
                _wordCellRange.Text = (i - 1).ToString();
            }

            for (int i = 0; i < list.Count; i++)
            {
                // № п/п
                _wordCellRange = _wordTable.Cell(i + 3, 1).Range;
                _wordCellRange.Text = list[i].Number.ToString();

                // Критерии оценки
                _wordCellRange = _wordTable.Cell(i + 3, 2).Range;
                _wordCellRange.Text = list[i].Name.ToString();

                // 2
                _wordCellRange = _wordTable.Cell(i + 3, 3).Range;
                _wordCellRange.Text = list[i].Mark2.ToString();

                // 3
                _wordCellRange = _wordTable.Cell(i + 3, 4).Range;
                _wordCellRange.Text = list[i].Mark3.ToString();

                // 4
                _wordCellRange = _wordTable.Cell(i + 3, 5).Range;
                _wordCellRange.Text = list[i].Mark4.ToString();

                // 5
                _wordCellRange = _wordTable.Cell(i + 3, 6).Range;
                _wordCellRange.Text = list[i].Mark5.ToString();

                // Краткое обоснование оценки
                _wordCellRange = _wordTable.Cell(i + 3, 7).Range;
                _wordCellRange.Text = list[i].Comment.ToString();
            }

            // Выравнивание
            begCell = _wordTable.Cell(1, 1).Range.Start;
            endCell = _wordTable.Cell(1, 4).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            _wordApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            begCell = _wordTable.Cell(2, 3).Range.Start;
            endCell = _wordTable.Cell(2, 6).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            _wordApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            begCell = _wordTable.Cell(3, 3).Range.Start;
            endCell = _wordTable.Cell(17, 6).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            _wordApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            begCell = _wordTable.Cell(3, 1).Range.Start;
            endCell = _wordTable.Cell(17, 1).Range.End;
            wordcellrange = _wordDocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            _wordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            _wordApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            #endregion

            // Меняем характеристики текста в таблице
            _wordTable.Range.Font.Color = Word.WdColor.wdColorBlack;
            _wordTable.Range.Font.Size = 12;
            _wordTable.Range.Font.Name = "Times New Roman";
            _wordTable.Range.Font.Italic = 0;
            _wordTable.Range.Font.Bold = 0;
        }

        /// <summary>
        /// Сохранение документа
        /// </summary>
        /// <param name="path">Путь для сохранения</param>
        public bool Save(string path)
        {
            try
            {
                _wordDocument.SaveAs($"{path}");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }

            return true;
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
