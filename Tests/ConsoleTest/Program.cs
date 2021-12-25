using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// ReSharper disable ArrangeMethodOrOperatorBody

namespace ConsoleTest
{
    class Program
    {
        private static void EditDocument(string data_file_name)
        {
            var data_file = new FileInfo(data_file_name);
            using var document = WordprocessingDocument.Open(data_file.FullName, true);

            var main_part = document.MainDocumentPart;
            var page = main_part.Document.Body;

            var styles = main_part.StyleDefinitionsPart.Styles.Elements<Style>().ToDictionary(s => s.StyleId);

            var changed = false;
            foreach (var paragraph in page.GetParagraphs())
            {
                if (paragraph.ParagraphProperties.ParagraphStyleId?.Val.Value is { } style_id)
                {
                    var style = styles[style_id];
                    Console.WriteLine(style.StyleName?.Val.Value);
                }

                Console.WriteLine("\t" + paragraph.InnerText);
                Console.WriteLine();

                if (paragraph.InnerText.Contains("qwe", StringComparison.OrdinalIgnoreCase))
                {
                    var text = paragraph.Descendants<Text>().First(t => t.InnerText.Contains("red", StringComparison.OrdinalIgnoreCase));
                    text.Text = text.Text.Replace("qwe", "red");
                    changed = true;
                }
            }

            if (changed)
                document.Save();
        }

        static void Main(string[] args)
        {
            var template = new FileInfo("Document.docx");
            var document = new FileInfo("doc.docx");
            document.Delete();
            template.CopyTo(document);

            var word = WordprocessingDocument.Open(document.FullName, true);
            var body = word.MainDocumentPart!.Document.Body!;

            foreach (var field in body.DescendantChilds<SdtElement>())
                field.ReplaceWithValue($"[{field.GetTag()}]:{field.InnerText}");

            //foreach (var block in body.DescendantChilds<SdtBlock>())
            //{
            //    var parent = block.Parent;
            //    var index = parent.FirstIndexOf(block);
            //    block.Remove();

            //    var run_properties = block.DescendantChilds<RunProperties>().First();
            //    run_properties.Remove();

            //    var run = new Run();
            //    run.AddChild(run_properties);
            //    run.AddChild(new Text("hEADER"));

            //    parent.InsertAt(new Paragraph(run), index);
            //}

            //var sdt_runs = body.DescendantChilds<SdtRun>()
            //   .ToLookup(r => r.GetTag());

            //var tags = sdt_runs.ToArray(r => r.Key);

            //foreach (var tag in tags)
            //{
            //    foreach (var run in sdt_runs[tag])
            //        run.ReplaceToRun($"({run.GetTag()}:{run.GetAlias()})[{run.GetText()}]~");

            //}

            //foreach (var sdt_run in body.DescendantChilds<SdtRun>()) 
            //    sdt_run.ReplaceToRun($"({sdt_run.GetTag()}:{sdt_run.GetAliase()})[{sdt_run.GetText()}]~");

            //word.SaveAs("test.docx");
            word.Close();

            document.Execute();

            //EditDocument("Document.docx");
            //CreateDocument("TestDoc.docx");
        }

        //private static IEnumerable<OpenXmlElement> EnumElements(OpenXmlElement element)
        //{
        //    yield return element;
        //    if (!element.HasChildren) yield break;

        //    foreach (var child_element in element.ChildElements)
        //        foreach (var child in EnumElements(child_element))
        //            yield return child;
        //}

        private static void CreateDocument(string FileName)
        {
            if (FileName is null) throw new ArgumentNullException(nameof(FileName));
            if (File.Exists(FileName)) File.Delete(FileName);

            using (var word_document = WordprocessingDocument.Create(FileName, WordprocessingDocumentType.Document))
            {
                var main_part = word_document.AddMainDocumentPart();

                var header_part = main_part.AddNewPart<HeaderPart>();
                header_part.Header = new()
                {
                    new Paragraph { "Test header!" }.Bold().AlignCenter()
                };
                var header_id = main_part.GetIdOfPart(header_part);

                var footer_part = main_part.AddNewPart<FooterPart>();
                footer_part.Footer = new()
                {
                    new Paragraph { "Test footer!" }.Bold().AlignRight()
                };
                var footer_id = main_part.GetIdOfPart(footer_part);

                var document = main_part.Document = new();
                document.Body = new()
                {
                    //new Paragraph(
                    //        new Run(new Text("111"))
                    //           .Bold()
                    //           .Color("FF0000"))
                    //   .Justification(JustificationValues.Center),
                    //new Paragraph(new Run(new Text("QWE")).Italic()),
                    //new Paragraph(new Run(new Text("ASD")).Underline()),
                    //new Paragraph { "QQQ" },
                    //table,
                    new Table
                    {
                        new TableRow
                        {
                            new TableCell { "123" }.Width(4672)
                               .Justification(JustificationValues.Center)
                               .Color("red")
                               .Border(6,0,6,6)
                               .VerticalAlignment(TableVerticalAlignmentValues.Center),
                            new TableCell { "qwe" }.Width(4672)
                               .Justification(JustificationValues.Center)
                               .Bold(),
                        }.Height(1009)
                    },
                    new SectionProperties()
                       .HeaderRef(header_id)
                       .FooterRef(footer_id)
                };
                word_document.Save();
            }

            Process.Start(new ProcessStartInfo(FileName) { UseShellExecute = true });
        }
    }
}
