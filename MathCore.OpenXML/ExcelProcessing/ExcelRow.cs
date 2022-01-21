#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MathCore.OpenXML.ExcelProcessing
{
    /// <summary>Строка таблицы</summary>
    /// <remarks>Осуществляет однопроходный доступ к своим ячейкам в процессе чтения документа</remarks>
    public readonly struct ExcelRow : IEnumerable<ExcelCell>, IEnumerable<string?>
    {
        /// <summary>Объект чтения должен быть сохранён до завершения перебора всех ячеек строки</summary>
        private readonly OpenXmlPartReader _Reader;
        /// <summary>Массив-таблица общих строк документа</summary>
        private readonly string[] _SharedStrings;
        private readonly int _Index;
        private readonly string? _Spans;
        private readonly int? _Style;
        private readonly int? _CustomFormat;
        private readonly double? _Height;
        private readonly double? _CustomHeight;
        private readonly bool _Collapsed;
        private readonly int _OutlineLevel;
        private readonly bool _Hidden;

        /// <summary>Индекс строки</summary>
        public int Index => _Index;

        /// <summary>Является ли строка скрытой</summary>
        public bool Hidden => _Hidden;

        /// <summary>Является ли строка сжатой</summary>
        public bool Collapsed => _Collapsed;

        /// <summary>Текстовые значения всех ячеек строки</summary>
        /// <remarks>При переборе значений все ячейки строки будут просмотрены и повторный доступ к ним будет невозможен</remarks>
        public IEnumerable<string?> Values
        {
            get
            {
                //return ((IEnumerable<ExcelCell>)this).Select(cell => cell.Value);
                var last_index = 0;
                foreach (var cell in this)
                {
                    var index_delta = cell.PositionIndex.Col - last_index;
                    for (var i = 1; i < index_delta; i++)
                        yield return null;
                    yield return cell.Value;
                    last_index += index_delta;
                }
            }
        }

        private static void ReadRowAttributes(
            IEnumerable<OpenXmlAttribute> Attributes,
            out int Index,
            out string? Spans,
            out int? Style,
            out int? CustomFormat,
            out double? Height,
            out double? CustomHeight,
            out bool Collapsed,
            out bool Hidden,
            out int OutlineLevel)
        {
            Index = default;
            Spans = default;
            Style = default;
            CustomFormat = default;
            Height = default;
            CustomHeight = default;
            Collapsed = default;
            Hidden = default;
            OutlineLevel = default;

            foreach (var attribute in Attributes)
                switch (attribute.LocalName)
                {
                    case "r":
                        Index = int.Parse(attribute.Value!);
                        break;
                    case "spans":
                        Spans = attribute.Value!;
                        break;
                    case "s":
                        Style = int.Parse(attribute.Value!);
                        break;
                    case "customFormat" when int.TryParse(attribute.Value, out var format):
                        CustomFormat = format;
                        break;
                    case "hidden":
                        Hidden = attribute.Value == "1";
                        break;
                    case "outlineLevel":
                        OutlineLevel = int.Parse(attribute.Value!);
                        break;
                    case "ht":
                        Height = double.Parse(attribute.Value!, CultureInfo.InvariantCulture);
                        break;
                    case "collapsed":
                        Collapsed = attribute.Value == "1";
                        break;
                    case "customHeight" when double.TryParse(attribute.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var height):
                        CustomHeight = height;
                        break;
                }
        }

        /// <summary>Инициализация нового экземпляра строки на основе объекта чтения данных файла</summary>
        /// <param name="Reader">Объект чтения данных файла</param>
        /// <param name="SharedStrings">Таблица общих строк файла</param>
        public ExcelRow(OpenXmlPartReader Reader, string[] SharedStrings)
        {
            if (Reader.ElementType != typeof(Row))
                throw new FormatException();

            ReadRowAttributes(Reader.Attributes,
                out _Index,
                out _Spans,
                out _Style,
                out _CustomFormat,
                out _Height,
                out _CustomHeight,
                out _Collapsed,
                out _Hidden,
                out _OutlineLevel);

            _Reader = Reader;
            _SharedStrings = SharedStrings;
        }

        IEnumerator<string?> IEnumerable<string?>.GetEnumerator() => Values.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public IEnumerator<ExcelCell> GetEnumerator()
        {
            if (_Reader.ElementType != typeof(Row) || !_Reader.IsStartElement)
                throw new InvalidOperationException("Некорректное состояние объекта чтения потока данных. Текущий элемент не является строкой.");

            var index = int.Parse(_Reader.Attributes.Value("r"));
            if (_Index != index)
                throw new InvalidOperationException("Попытка повторного перечисления ячеек строки невозможно");

            while (_Reader.Read())
                if (_Reader.ElementType == typeof(Cell) && _Reader.IsStartElement)
                    yield return new ExcelCell(_Reader, _SharedStrings);
                else if (_Reader.ElementType == typeof(Row))
                    break;
                else
                    _Reader.Skip();
        }
    }
}