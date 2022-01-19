using System;
using System.Collections.Generic;

namespace MathCore.OpenXML.WordProcessing.Templates;

public interface IFieldValueSetter
{
    object this[string FieldName] { set; }

    IFieldValueSetter Field(string FieldName, string Value);
    IFieldValueSetter Field(string FieldName, Func<string> Value);
    IFieldValueSetter Field(string FieldName, object Value);
    IFieldValueSetter Field<T>(string FieldName, T Value);

    void SetValue(string Value);
    void SetValue(Func<string> Value);
    void SetValue(object Value);

    IFieldValueSetter Field<T>(string FieldName, IEnumerable<T> Values, Action<IFieldValueSetter, T> Setter);
}