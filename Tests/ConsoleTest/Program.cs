using System.Globalization;
using ConsoleTest.Models;
using MathCore.OpenXML.WordProcessing;

var template_file = new FileInfo("Document.docx");
var document_file = new FileInfo("doc.docx");

var products = Product.Test(rnd: new Random()).ToArray();

products[0] = products[0] with { Features = [] };

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
