using System;
using System.Collections.Generic;
using System.Linq;

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
        private Dictionary<string, SdtElement> _Fields = null!;
        private bool _ReplaceFieldsWithValues;

        public object this[string FieldName]
        {
            set
            {
                switch (value)
                {
                    default: Field(FieldName, value); break;
                    case string v: Field(FieldName, v); break;
                    case Func<string> v: Field(FieldName, v); break;
                }
            }
        }

        public FieldValueSetter(Action<IFieldValueSetter, T> Setter) => _Setter = Setter;

        private void SetField(string FieldName, string Value)
        {
            if (Value is null) return;

            if (_Fields is not { Count: > 0 } fields || !fields.TryGetValue(FieldName, out var field)) return;


            if (_ReplaceFieldsWithValues)
                field.ReplaceWithContentValue(Value);
            else
                field.SetContentValue(Value);
        }

        public IFieldValueSetter Field(string FieldName, string Value)
        {
            SetField(FieldName, Value);
            return this;
        }

        public IFieldValueSetter Field(string FieldName, Func<string> Value) => Field(FieldName, Value());

        public IFieldValueSetter Field(string FieldName, object Value) => Field(FieldName, Value.ToString());

        public void SetValue(string Value)
        {
            var paragraph = _CurrentElement as Paragraph
                ?? _CurrentElement.Descendants<Paragraph>().First();
            paragraph.Text(Value);
        }

        public void SetValue(Func<string> Value) => SetValue(Value());

        public void SetValue(object Value) => SetValue(Value.ToString());

        public IFieldValueSetter FieldEnum<TValue>(
            string FieldName,
            IEnumerable<TValue> Values,
            Action<IFieldValueSetter, TValue> Setter)
        {
            if (_Fields is not { Count: > 0 } fields || !fields.TryGetValue(FieldName, out var field))
                return this;

            var block = Create(FieldName, Values, Setter);
            block.Process(field, _ReplaceFieldsWithValues);

            return this;
        }

        public OpenXmlElement CreateElement(T Value, OpenXmlElement Template, bool ReplaceFieldsWithValues)
        {
            var element = Template.CloneNode(true);
            _CurrentElement = element;
            _ReplaceFieldsWithValues = ReplaceFieldsWithValues;
            _Fields = element.Descendants<SdtElement>()
               .Select(Element => (Tag: Element.GetTag(), Element))
               .Where(e => e.Tag is { Length: > 0 })
               .ToDictionary(e => e.Tag, e => e.Element)!;

            _Setter(this, Value);

            return element;
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
            Process(field, ReplaceFieldsWithValues);
    }

    private void Process(SdtElement Field, bool ReplaceFieldsWithValues)
    {
        var field = Field;
        if (field is SdtCell sdt_cell)
        {
            var cell = sdt_cell.ReplaceWithContent();
            field = cell.GetFirstChild<SdtBlock>() 
                ?? throw new InvalidOperationException("Не найден шаблонный блок в шаблонной ячейке таблицы");
        }

        var parent = field.Parent ?? throw new InvalidOperationException("Не найден родительский элемент");
        var index = parent!.FirstIndexOf(field);
        field.Remove();

        var template = field.GetContent();
        foreach (var value in _Values)
        {
            var element = _ValueSetter.CreateElement(value, template, ReplaceFieldsWithValues);
            parent.InsertAt(element, index);
            index++;
        }
    }
}