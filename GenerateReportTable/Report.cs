using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateReportTable
{
    /// <summary>
    /// Класс для создания оценки для отчета
    /// </summary>
    class Report
    {
        private int _number;
        private string _name;
        private char _mark2;
        private char _mark3;
        private char _mark4;
        private char _mark5;
        private string _comment;

        /// <summary>
        /// Номер по порядку
        /// </summary>
        public int Number { get => _number; set => _number = value; }
        /// <summary>
        /// Критерий оценки
        /// </summary>
        public string Name { get => _name; set => _name = value; }
        /// <summary>
        /// 2
        /// </summary>
        public char Mark2 { get => _mark2; set => _mark2 = value; }
        /// <summary>
        /// 3
        /// </summary>
        public char Mark3 { get => _mark3; set => _mark3 = value; }
        /// <summary>
        /// 4
        /// </summary>
        public char Mark4 { get => _mark4; set => _mark4 = value; }
        /// <summary>
        /// 5
        /// </summary>
        public char Mark5 { get => _mark5; set => _mark5 = value; }
        /// <summary>
        /// Краткое обоснование оценки
        /// </summary>
        public string Comment { get => _comment; set => _comment = value; }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="number">Номер по порядку</param>
        /// <param name="mark">Оценка</param>
        public Report(int number, int mark)
        {
            _number = number;

            // Заполняем критерий оценки
            switch (number)
            {
                case 1:
                    _name = "Полнота и правильность раскрытия темы";
                    break;
                case 2:
                    _name = "Логическое и последовательное изложение темы";
                    break;
                case 3:
                    _name = "Характер изложения материала";
                    break;
                case 4:
                    _name = "Стиль и убедительность изложения";
                    break;
                case 5:
                    _name = "Умение укладываться в отведенное время";
                    break;
                case 6:
                    _name = "Темп речи";
                    break;
                case 7:
                    _name = "Использование специально подготовленных иллюстративных материалов (презентации)";
                    break;
                case 8:
                    _name = "Уверенность и спокойствие выступающего";
                    break;
                case 9:
                    _name = "Грамотность, выразительность речи, дикция";
                    break;
                case 10:
                    _name = "Жестикуляция";
                    break;
                case 11:
                    _name = "Ошибки и оговорки во время выступления";
                    break;
                case 12:
                    _name = "Общая манера поведения выступающего";
                    break;
                case 13:
                    _name = "Собственное отношение к излагаемой проблеме";
                    break;
                case 14:
                    _name = "Уровень обратной связи";
                    break;
                case 15:
                    _name = "Общая оценка";
                    break;
            }

            // Генерируем маркер
            Random rnd = new Random();
            int m = rnd.Next(0, 3);
            char marker;

            switch (m)
            {
                case 0:
                    marker = '+';
                    break;
                case 1:
                    marker = 'v';
                    break;
                case 2:
                    marker = 'x';
                    break;
                default:
                    marker = '+';
                    break;
            }

            // Проверяем, чтобы не была строка Общая оценка
            if (number != 15)
            {
                // Заполняем оценки
                switch (mark)
                {
                    case 2:
                        _mark2 = marker;
                        _mark3 = ' ';
                        _mark4 = ' ';
                        _mark5 = ' ';
                        break;
                    case 3:
                        _mark2 = ' ';
                        _mark3 = marker;
                        _mark4 = ' ';
                        _mark5 = ' ';
                        break;
                    case 4:
                        _mark2 = ' ';
                        _mark3 = ' ';
                        _mark4 = marker;
                        _mark5 = ' ';
                        break;
                    case 5:
                        _mark2 = ' ';
                        _mark3 = ' ';
                        _mark4 = ' ';
                        _mark5 = marker;
                        break;
                }
            }
            else
            {
                _mark2 = ' ';
                _mark3 = ' ';
                _mark4 = marker;
                _mark5 = ' ';
            }
            
        }
    }
}
