using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing.Templates;

public class TemplateFieldValue : TemplateField
{
    private readonly object _Value;

    public string Value => _Value switch
    {
        string str => str,
        Func<string> f => f(),
        _ => _Value.ToString()
    };

    public TemplateFieldValue(string Tag, object Value) : base(Tag) => _Value = Value;

    public TemplateFieldValue(string Tag, string Value) : base(Tag) => _Value = Value;

    public TemplateFieldValue(string Tag, Func<string> Value)
        : base(Tag) => _Value = Value;

    public override void Process(IEnumerable<SdtElement> Fields, bool ReplaceFieldsWithValues)
    {
        var value = Value;
        if (ReplaceFieldsWithValues)
            foreach (var field in Fields)
                field.ReplaceWithContentValue(value);
        else
            foreach (var field in Fields)
                field.SetContentValue(value);
    }
}

public class TemplateFieldValue<T> : TemplateField
{
    private readonly T _Value;

    public TemplateFieldValue(string Tag, T Value) : base(Tag) => _Value = Value;

    public override void Process(IEnumerable<SdtElement> Fields, bool ReplaceFieldsWithValues)
    {
        var value = _Value as string ?? (_Value as Func<string>)?.Invoke() ?? _Value?.ToString();
        if(value is null) return;

        if (ReplaceFieldsWithValues)
            foreach (var field in Fields)
                field.ReplaceWithContentValue(value);
        else
            foreach (var field in Fields)
                field.SetContentValue(value);
    }
}