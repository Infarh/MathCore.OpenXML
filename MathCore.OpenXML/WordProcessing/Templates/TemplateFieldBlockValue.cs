using System;
using System.Collections.Generic;
using System.Linq;
using System.Tags;
using System.Xml.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing.Templates;

public abstract class TemplateFieldBlockValue : TemplateField
{
    public static TemplateFieldBlockValue<T> Create<T>(
        string TagName, IEnumerable<T> Values,
        Action<IFieldValueSetter, T> Setter) =>
        new(TagName, Values, Setter);

    protected TemplateFieldBlockValue(string Tag) : base(Tag) { }
}

public class TemplateFieldBlockValue<T> : TemplateFieldBlockValue
{
    private class FieldValueSetter : IFieldValueSetter
    {
        private readonly Action<IFieldValueSetter, T> _Setter;
        private OpenXmlElement _CurrentElement = null!;
        private ILookup<string?, SdtElement> _Fields = null!;
        private bool _ReplaceFieldsWithValues;

        public object this[string FieldName]
        {
            set
            {
                switch (value)
                {
                    default:
                        Field(FieldName, value);
                        break;
                    case string v:
                        Field(FieldName, v);
                        break;
                    case Func<string> v:
                        Field(FieldName, v);
                        break;
                }
            }
        }

        public FieldValueSetter(Action<IFieldValueSetter, T> Setter) => _Setter = Setter;

        private void SetField(string FieldName, string Value)
        {
            if (Value is null) return;

            if (_Fields.Count == 0) return;

            if (_ReplaceFieldsWithValues)
                foreach (var field in _Fields[FieldName])
                    field.ReplaceWithContentValue(Value);
            else
                foreach (var field in _Fields[FieldName])
                    field.SetContentValue(Value);
        }

        public IFieldValueSetter Field(string FieldName, string Value)
        {
            SetField(FieldName, Value);
            return this;
        }

        public IFieldValueSetter Field(string FieldName, Func<string> Value) => Field(FieldName, Value());

        public IFieldValueSetter Field(string FieldName, object Value) => Field(FieldName, Value.ToString());
        public IFieldValueSetter Field<TValue>(string FieldName, TValue Value) => Value is null ? this : Field(FieldName, Value.ToString());

        public void SetValue(string Value)
        {
            if (_CurrentElement is null)
                throw new NotSupportedException("Добавление текстового значения в форму невозможно. Требуется указать название поля.");

            var paragraph = _CurrentElement as Paragraph
                ?? _CurrentElement.Descendants<Paragraph>().First();
            paragraph.Text(Value);
        }

        public void SetValue(Func<string> Value) => SetValue(Value());

        public void SetValue(object Value) => SetValue(Value.ToString());

        public IFieldValueSetter Field<TValue>(
            string FieldName,
            IEnumerable<TValue> Values,
            Action<IFieldValueSetter, TValue> Setter)
        {
            if (_Fields.Count == 0)
                return this;

            var fields = _Fields[FieldName];

            var values = Values.ToArray();
            foreach (var field in fields)
            {
                var block = Create(FieldName, values, Setter);
                block.ProcessField(field, _ReplaceFieldsWithValues);
            }

            return this;
        }

        public OpenXmlElement FeelElement(T Value, OpenXmlElement Element, bool ReplaceFieldsWithValues)
        {
            _CurrentElement = Element;
            _ReplaceFieldsWithValues = ReplaceFieldsWithValues;

            _Fields = Element.GetFields()
               .Select(e => (Tag: e.GetTag(), Element: e))
               .Where(e => e.Tag is { Length: > 0 })
               .ToLookup(e => e.Tag, e => e.Element);

            _Setter(this, Value);

            return Element;
        }

        public void FeelElements(T Value, IReadOnlyList<OpenXmlElement> Elements, bool ReplaceFieldsWithValues)
        {
            _ReplaceFieldsWithValues = ReplaceFieldsWithValues;

            _Fields = Elements.SelectMany(e => e.GetFields())
               .Select(e => (Tag: e.GetTag(), Element: e))
               .Where(e => e.Tag is { Length: > 0 })
               .ToLookup(e => e.Tag, e => e.Element);

            _Setter(this, Value);
        }
    }

    private readonly IEnumerable<T> _Values;
    private readonly FieldValueSetter _ValueSetter;

    public TemplateFieldBlockValue(string Tag, IEnumerable<T> Values, Action<IFieldValueSetter, T> Setter)
        : base(Tag)
    {
        _Values = Values;
        _ValueSetter = new(Setter);
    }

    public override void Process(IEnumerable<SdtElement> Fields, bool ReplaceFieldsWithValues)
    {
        foreach (var field in Fields)
            ProcessField(field, ReplaceFieldsWithValues);
    }

    private void ProcessField(SdtElement Field, bool ReplaceFieldsWithValues)
    {
        switch (Field)
        {
            case SdtBlock block:
                ProcessBlock(block, ReplaceFieldsWithValues);
                break;

            case SdtCell cell:
                ProcessCell(cell, ReplaceFieldsWithValues);
                break;

            default:
                Process(Field, ReplaceFieldsWithValues);
                break;
        }
    }

    private void Process(SdtElement Field, bool ReplaceFieldsWithValues)
    {
        OpenXmlElement last_element = Field;
        var template = Field.GetContent();
        foreach (var value in _Values)
        {
            last_element = last_element.InsertAfterSelf(template.CloneNode(true));
            _ValueSetter.FeelElement(value, last_element, ReplaceFieldsWithValues);
        }

        Field.Remove();
    }

    private void ProcessCell(SdtCell CellField, bool ReplaceFieldsWithValues)
    {
        var cell = CellField.ReplaceWithContent();
        var field = cell.GetFirstChild<SdtBlock>()
            ?? throw new InvalidOperationException("Не найден шаблонный блок в шаблонной ячейке таблицы");

        OpenXmlElement last_element = field;
        var template = field.GetContent();
        foreach (var value in _Values)
        {
            last_element = last_element.InsertAfterSelf(template.CloneNode(true));
            _ValueSetter.FeelElement(value, last_element, ReplaceFieldsWithValues);
        }

        field.Remove();
    }

    private void ProcessBlock(SdtBlock BlockField, bool ReplaceFieldsWithValues)
    {
        var block_content_0 = BlockField.SdtContentBlock?.FirstChild as SdtBlock ?? throw new InvalidOperationException("Содержимое шаблонного блока не найдено");
        var template_elements = block_content_0.SdtContentBlock?.ChildElements ?? throw new InvalidOperationException("Дочерние элементы шаблонного блока не определены.");

        //OpenXmlElement last_element = BlockField;
        var elements = new List<OpenXmlElement>(template_elements.Count);
        foreach (var value in _Values)
        {
            elements.AddRange(template_elements.Select(e => e.CloneNode(true)));
            BlockField.InsertBeforeSelf(elements[0]);
            for (var i = 1; i < elements.Count; i++) 
                elements[i - 1].InsertAfterSelf(elements[i]);

            _ValueSetter.FeelElements(value, elements, ReplaceFieldsWithValues);

            elements.Clear();
        }

        BlockField.Remove();
    }
}