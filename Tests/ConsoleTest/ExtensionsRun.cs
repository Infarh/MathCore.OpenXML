using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleTest
{
    public static class ExtensionsRun
    {
        public static Run SetProperties(this Run run, RunProperties properties)
        {
            run.RunProperties = properties;
            return run;
        }

        public static Run Bold(this Run run, bool IsBold = true)
        {
            var properties = run.RunProperties ??= new();
            properties.Bold = IsBold ? new() : null;
            return run;
        }

        public static Run Italic(this Run run, bool IsItalic = true)
        {
            var properties = run.RunProperties ??= new();
            properties.Italic = IsItalic ? new() : null;
            return run;
        }

        public static Run Underline(this Run run, bool IsUnderline = true)
        {
            var properties = run.RunProperties ??= new();
            properties.Underline = IsUnderline ? new() : null;
            return run;
        }

        public static Run Color(this Run run, string Color)
        {
            var properties = run.RunProperties ??= new();
            var color = properties.Color ??= new();
            color.Val = Color;
            return run;
        }

        public static Run FontSize(this Run run, int Size)
        {
            var properties = run.RunProperties ??= new();
            var font_size = properties.FontSize ??= new();
            font_size.Val = Size.ToString();
            return run;
        }

        public static Run Font(this Run run, string FontName)
        {
            var properties = run.RunProperties ??= new();
            var run_fonts = properties.RunFonts ??= new();
            run_fonts.Ascii = FontName;
            run_fonts.HighAnsi = FontName;
            return run;
        }

        /// <summary>Разрежение символов</summary>
        public static Run Spacing(this Run run, int Size)
        {
            var properties = run.RunProperties ??= new();
            var spacing = properties.Spacing ??= new();
            spacing.Val = Int32Value.FromInt32(Size);
            return run;
        }

        public static Run Language(this Run run, string Language)
        {
            var properties = run.RunProperties ??= new();
            if(string.IsNullOrEmpty(Language))
                properties.RemoveAllChildren<Languages>();
            else
            {
                var languages = properties.Languages ??= new();
                languages.Val = Language;
            }

            return run;
        }

        public static Run AppendText(this Run run, string str)
        {
            run.AppendChild(new Text { Text = str });
            return run;
        }

        public static Run Text(this Run run, string str)
        {
            if (run.GetFirstChild<Text>() is { } text) 
                text.Text = str;
            else
                run.AppendChild(new Text(str));

            return run;
        }

        public static Run Tab(this Run run)
        {
            run.AppendChild(new TabChar());
            return run;
        }
    }
}
