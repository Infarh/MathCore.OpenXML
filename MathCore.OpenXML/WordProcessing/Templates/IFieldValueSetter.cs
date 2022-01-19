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

    void Value(string Value);
    void Value(Func<string> Value);
    void Value(object Value);

    IFieldValueSetter Field<T>(string FieldName, IEnumerable<T> Values, Action<IFieldValueSetter, T> Setter);
}