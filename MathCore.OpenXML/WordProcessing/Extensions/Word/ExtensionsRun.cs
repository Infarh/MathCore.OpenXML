using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsRun
{
    extension(Run)
    {
        public static Run WithText(string Text) => new(new Text(Text));

        public static Run WithText(string Text, string StyleId) => new(
            new RunProperties(new RunStyle { Val = StyleId }),
            new Text(Text));
    }

    extension(Run run)
    {
        public Run SetProperties(RunProperties properties)
        {
            run.RunProperties = properties;
            return run;
        }

        public Run Bold(bool IsBold = true)
        {
            var properties = run.RunProperties ??= new();
            properties.Bold = IsBold ? new() : null;
            return run;
        }

        public Run Italic(bool IsItalic = true)
        {
            var properties = run.RunProperties ??= new();
            properties.Italic = IsItalic ? new() : null;
            return run;
        }

        public Run Underline(bool IsUnderline = true)
        {
            var properties = run.RunProperties ??= new();
            properties.Underline = IsUnderline ? new() : null;
            return run;
        }

        public Run Color(string Color)
        {
            var properties = run.RunProperties ??= new();
            var color = properties.Color ??= new();
            color.Val = Color;
            return run;
        }

        public Run FontSize(int Size)
        {
            var properties = run.RunProperties ??= new();
            var font_size = properties.FontSize ??= new();
            font_size.Val = Size.ToString();
            return run;
        }

        public Run Font(string FontName)
        {
            var properties = run.RunProperties ??= new();
            var run_fonts = properties.RunFonts ??= new();
            run_fonts.Ascii = FontName;
            run_fonts.HighAnsi = FontName;
            return run;
        }

        /// <summary>Разрежение символов</summary>
        public Run Spacing(int Size)
        {
            var properties = run.RunProperties ??= new();
            var spacing = properties.Spacing ??= new();
            spacing.Val = Size;
            return run;
        }

        public Run Language(string Language)
        {
            var properties = run.RunProperties ??= new();
            if (string.IsNullOrEmpty(Language))
                properties.RemoveAllChildren<Languages>();
            else
            {
                var languages = properties.Languages ??= new();
                languages.Val = Language;
            }

            return run;
        }

        public Run AppendText(string str)
        {
            run.AppendChild(new Text { Text = str });
            return run;
        }

        public Run Text(string str)
        {
            if (run.GetFirstChild<Text>() is { } text)
                text.Text = str;
            else
                run.AppendChild(new Text(str));

            return run;
        }

        public Run Tab()
        {
            run.AppendChild(new TabChar());
            return run;
        }

        public Run Style(string StyleId)
        {
            var properties = run.RunProperties ??= new();
            properties.RunStyle = new RunStyle { Val = StyleId };
            return run;
        }
    }
}