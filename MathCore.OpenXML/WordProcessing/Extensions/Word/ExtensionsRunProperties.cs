using DocumentFormat.OpenXml.Wordprocessing;

namespace MathCore.OpenXML.WordProcessing.Extensions.Word;

public static class ExtensionsRunProperties
{
    extension(RunProperties)
    {
        public static RunProperties withStyle(string StyleId) => new(new RunStyle { Val = StyleId });
    }

    extension(RunProperties properties)
    {
        public RunProperties Style(string StyleId)
        {
            properties.RunStyle = new RunStyle { Val = StyleId };
            return properties;
        }
    }
}
