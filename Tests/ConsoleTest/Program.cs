using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

using ConsoleTest.Models;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using MathCore.OpenXML.WordProcessing;
using MathCore.OpenXML.WordProcessing.Extensions.Word;

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
            ExcelWriterTest.Run();

            return;

            var template_file = new FileInfo("Document.docx");
            var document_file = new FileInfo("doc.docx");

            var products = Product.Test(rnd: new Random()).ToArray();

            products[0] = products[0] with { Features = Enumerable.Empty<Product.Feature>() };

            var template = Word.Template(template_file)
                   .Field("ProductCart", products, (product_cart, product) => product_cart
                       .Field("Name", product.Name)
                       .Field("Id", product.Id)
                       .Field("Price", product.Price.ToString("C2"))
                       .Field("Description", product.Description)
                       .Field("Feature", product.Features, (feature_cart, feature) => feature_cart
                           .Field("Name", feature.Name)
                           .Field("Id", feature.Id)
                           .Field("Description", feature.Description)))
                   .Field("CatalogName", "Компьютеры")
                   .Field("CreationTime", DateTime.Now.ToString("f", CultureInfo.GetCultureInfo("ru")))
                   .Field("ProductsCount", products.Length)
                   .Field("ProductTotalPrice", products.DefaultIfEmpty().Sum(p => p?.Price ?? 0))
                   .Field("ProductInfo", products, (product_row, product) => product_row
                       .Field("ProductId", product.Id)
                       .Field("ProductName", product.Name)
                       .Field("ProductPrice", product.Price)
                       .Field("ProductFeature", product.Features, (feature_row, feature) => feature_row
                           .Value(feature.Description)))
                   .RemoveUnprocessedFields()
                   .ReplaceFieldsWithValues()
                ;

            var file = template.SaveTo(document_file);
            file.Execute();

            //EditDocument("Document.docx");
            //CreateDocument("TestDoc.docx");
        }

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
                    new Paragraph { "Test footer!" }
                       .Bold()
                       .AlignRight()
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
