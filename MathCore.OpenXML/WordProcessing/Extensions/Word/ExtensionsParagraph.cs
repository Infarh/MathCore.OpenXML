using System;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsParagraph
{
    public static T Justification<T>(this T element, JustificationValues? Justification) where T : OpenXmlElement
    {
        foreach (var run in element.Descendants<Paragraph>())
            run.Justification(Justification);
        return element;
    }

    public static T AlignLeft<T>(this T element) where T : OpenXmlElement => element.Justification(JustificationValues.Left);
    public static T AlignRight<T>(this T element) where T : OpenXmlElement => element.Justification(JustificationValues.Right);
    public static T AlignCenter<T>(this T element) where T : OpenXmlElement => element.Justification(JustificationValues.Center);

    public static Paragraph Add(this Paragraph paragraph, Run run)
    {
        paragraph.AppendChild(run);
        return paragraph;
    }

    public static Paragraph Add(this Paragraph paragraph, string text)
    {
        paragraph.Add(new Run(new Text(text)));
        return paragraph;
    }

    public static Paragraph Justification(this Paragraph paragraph, JustificationValues? Justification)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        properties.Justification = Justification is null ? null : new Justification { Val = Justification };

        return paragraph;
    }

    public static Paragraph AlignLeft(this Paragraph paragraph) => paragraph.Justification(JustificationValues.Left);
    public static Paragraph AlignRight(this Paragraph paragraph) => paragraph.Justification(JustificationValues.Right);
    public static Paragraph AlignCenter(this Paragraph paragraph) => paragraph.Justification(JustificationValues.Center);

    public static Paragraph SpacingBetweenLines(this Paragraph paragraph,
        int Before,
        int After = 0,
        int Line = 240,
        LineSpacingRuleValues? LineRile = null)
    {
        var spacing = (paragraph.ParagraphProperties ??= new()).SpacingBetweenLines ??= new();

        // Множитель 20: чтобы установить 18 надо задать 360
        if (Before > 0) spacing.Before = Before.ToString();
        if (After > 0) spacing.After = After.ToString();
        if (Line > 0) spacing.Line = Line.ToString();
        spacing.LineRule = LineRile ?? LineSpacingRuleValues.Auto;

        return paragraph;
    }

    public static Paragraph Bold(this Paragraph paragraph, bool IsBold = true)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        properties.ParagraphMarkRunProperties ??= new();

        if (IsBold)
            properties.Content().GetOrAppend<Bold>();
        else
            properties.RemoveAllChildren<Bold>();

        return paragraph;
    }

    public static Paragraph Italic(this Paragraph paragraph, bool IsItalic = true)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        properties.ParagraphMarkRunProperties ??= new();

        if (IsItalic)
            properties.Content().GetOrAppend<Italic>();
        else
            properties.RemoveAllChildren<Italic>();

        return paragraph;
    }

    public static Paragraph Underline(this Paragraph paragraph, bool IsUnderline = true)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        properties.ParagraphMarkRunProperties ??= new();

        if (IsUnderline)
            properties.Content().GetOrAppend<Underline>();
        else
            properties.RemoveAllChildren<Underline>();

        return paragraph;
    }

    public static Paragraph Color(this Paragraph paragraph, string Color)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        properties.ParagraphMarkRunProperties ??= new();

        if (string.IsNullOrEmpty(Color))
            properties.RemoveAllChildren<Color>();
        else
            properties.Content().GetOrAppend<Color>().Val = Color;

        return paragraph;
    }

    public static Paragraph FontSize(this Paragraph paragraph, int Size)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        properties.ParagraphMarkRunProperties ??= new();

        if (Size <= 0)
            properties.RemoveAllChildren<FontSize>();
        else
            properties.Content().GetOrAppend<FontSize>().Val = Size.ToString();

        return paragraph;
    }

    public static Paragraph Font(this Paragraph paragraph, string FontName)
    {
        var properties = (paragraph.ParagraphProperties ??= new())
           .ParagraphMarkRunProperties ??= new();

        if (string.IsNullOrEmpty(FontName))
            properties.RemoveAllChildren<RunFonts>();
        else
        {
            var font = properties.Content().GetOrAppend<RunFonts>();
            font.Ascii = font.HighAnsi = FontName;
        }

        return paragraph;
    }

    public static Paragraph Language(this Paragraph paragraph, string Language)
    {
        var properties = (paragraph.ParagraphProperties ??= new())
           .ParagraphMarkRunProperties ??= new();

        if (string.IsNullOrEmpty(Language))
            properties.RemoveAllChildren<Languages>();
        else
        {
            var language = properties.Content().GetOrAppend<Languages>();
            language.Val = Language;
        }

        return paragraph;
    }

    public static Paragraph TabStop(this Paragraph paragraph, int Position, TabStopValues Type)
    {
        var properties = paragraph.ParagraphProperties ??= new();
        var tabs = properties.Tabs ??= new();

        tabs.AppendChild(new TabStop
        {
            Position = Position,
            Val = Type,
        });

        return paragraph;
    }

    public static Run Text(this Paragraph paragraph, string Text)
    {
        var runs = paragraph.Descendants<Run>().ToArray();
        if (runs.Length == 0) 
            return paragraph.AppendChild(new Run().Text(Text));

        runs[0].Text(Text);
        for(var i = 1; i < runs.Length; i++)
            runs[i].Remove();

        return runs[0];
    }
}