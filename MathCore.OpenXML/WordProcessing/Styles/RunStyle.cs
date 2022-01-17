using System.ComponentModel;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Styles
{
    public class RunStyle
    {
        public string Font { get; set; } = "Times New Roman";

        public int FontSize { get; set; } = 28;

        public bool Bold { get; set; }

        public enum ThemeType
        {
            MajorEastAsia,
            MajorBidi,
            MajorAscii,
            MajorHighAnsi,
            MinorEastAsia,
            MinorBidi,
            MinorAscii,
            MinorHighAnsi,
        }

        public ThemeType? EastAsiaTheme { get; set; }

        public ThemeType? ComplexScriptTheme { get; set; }

        private RunFonts CreateRunFonts()
        {
            var properties = new RunFonts();

            if (Font is { Length: > 0 } font)
            {
                properties.Ascii = font;
                properties.HighAnsi = font;
            }

            if (EastAsiaTheme is { } east_asia_theme)
                properties.EastAsiaTheme = east_asia_theme switch
                {
                    ThemeType.MajorEastAsia => ThemeFontValues.MajorEastAsia,
                    ThemeType.MajorBidi => ThemeFontValues.MajorBidi,
                    ThemeType.MajorAscii => ThemeFontValues.MajorAscii,
                    ThemeType.MajorHighAnsi => ThemeFontValues.MajorHighAnsi,
                    ThemeType.MinorEastAsia => ThemeFontValues.MinorEastAsia,
                    ThemeType.MinorBidi => ThemeFontValues.MinorBidi,
                    ThemeType.MinorAscii => ThemeFontValues.MinorAscii,
                    ThemeType.MinorHighAnsi => ThemeFontValues.MinorHighAnsi,
                    _ => throw new InvalidEnumArgumentException(nameof(EastAsiaTheme), (int)east_asia_theme, typeof(ThemeType))
                };

            if (ComplexScriptTheme is { } complex_script_theme)
                properties.EastAsiaTheme = complex_script_theme switch
                {
                    ThemeType.MajorEastAsia => ThemeFontValues.MajorEastAsia,
                    ThemeType.MajorBidi => ThemeFontValues.MajorBidi,
                    ThemeType.MajorAscii => ThemeFontValues.MajorAscii,
                    ThemeType.MajorHighAnsi => ThemeFontValues.MajorHighAnsi,
                    ThemeType.MinorEastAsia => ThemeFontValues.MinorEastAsia,
                    ThemeType.MinorBidi => ThemeFontValues.MinorBidi,
                    ThemeType.MinorAscii => ThemeFontValues.MinorAscii,
                    ThemeType.MinorHighAnsi => ThemeFontValues.MinorHighAnsi,
                    _ => throw new InvalidEnumArgumentException(nameof(EastAsiaTheme), (int)complex_script_theme, typeof(ThemeType))
                };


            return properties;
        }

        public StyleRunProperties CreateProperties()
        {
            var properties = new StyleRunProperties(CreateRunFonts());

            if (Bold) properties.Append(new Bold());

            properties.Append(new FontSize { Val = FontSize.ToString() });
            properties.Append(new FontSizeComplexScript { Val = FontSize.ToString() });

            return properties;
        }
    }
}
