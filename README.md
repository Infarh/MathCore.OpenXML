# Инструменты для работы с документами OpenXML

Библиотека предназначени для управления документами в формате OpenXML
- чтение
- редактирование
- генерация
- заполнение форм отчётов

  ## Установка

  https://www.nuget.org/packages/MathCore.OpenXML

  ```Shell
  dotnet add package MathCore.OpenXML
  ```

  ## Заполнение word-шаблона

Для создания шаблона документа требуется
- создать обычный Word-документ
- Заполнить базовую часть шаблона (написать текст, добавить таблицы и т.п.)
- открыть меню разработчика
  - Параметры/Настроить ленту
  - Включить раздел "Разработчик"
- В местах, куда требуется вставка данных в шаблоне создать "поля"
  - Выделить текст под замену
  - На вкладке "Разработчик" нажать кнопку "Aa" - "Элемент упралвения содержимым - Обычный текст"
  - Для созданного поля открыть его свойства - на панели "Разработчик" нажать кнопку "СВойства"
  - В диалоге настройки поля указать его идентификатор "Тег"
  - Идентификаторы не обязательно должны быть уникальными.
  - Для таблиц поля должны быть расставлены в заполняемой строке, а также вся строка должна быть назначена полем.
 
После подготовки шаблона в годе заполнение выполняется следующим образом

```C#
var template_file = new FileInfo("Document.docx")
  
var template = Word.Template(template_file) // подготовка шаблона - файл не открывается и результат не формируется на данном этапе
     .Field("ProductCart", products, (product_cart, product) => product_cart // заполнение комплексного поля, содержащего внутри себя шаблоны
         .Field("Name", product.Name) // заполнение простого поля
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
     .Field("FooterFileInfo", "Смета №1")
     .RemoveUnprocessedFields() // указание, что надо удалить необработанные поля
     .ReplaceFieldsWithValues() // указание, что надо заменить все поля их исходным текстом, если поля не были обработаны
  ;
  ```

После заполнения шаблона его надо сохранить

```C#
FileInfo report_file = template.SaveTo(document_file);
```

либо так

```C#
using (var stream = document_file.Create())
    template.SaveTo(stream);
```

Также шаблон позволяет перечислить все поля и получить доступ к их содержимому

```C#
var fields = template.EnumerateFields().ToArray();
```

## Чтение данных xlsx-файлов

```C#
// буфер чтения содержимого ячеек очередной строки.
// Ожидаем что в строке будет порядка 100 ячеек. Если не угадаем, то ничего страшного.
var row = new List<string>(100);
var row_index = 1; // номер строки. Начинаем с 1 так как одну строку пропустим (заголовок)
foreach(var data in Excel.File("file.xslx")["SheetName"].Skip(1))
{
  row_index++; // сразу инкрементируем номер строки
  row.Clear(); // чистим буфер чтения
  row.AddRange(data); // загружаем все ячейки в буфер

  // далее анализируем содержимое буфера. Ячейки могут содержать как пустые строки, так и null
  // Можно это сделать так:
  if(row is not [
      var id_str,                 // первая ячейка будет загружена в переменную id_str в любом случае
      { Length: > 0 } name,   // вторая ячейка должна содержать строку ненулевой длины
      ['D', ..] description,  // третья ячейка должна иметь строку, начинающуюся с символа D
      .. // могут быть ещё ячейки
    ])
  contunue; // если строка не соответствует шаблону, то переходим к следующей

  if(!int.TrypParse(id_str, out var id))
    throw new InvalidOperationException($"Ошибка формата файла. Идентификатор в строке {row_index} не является числом. Содержимое ячейки {id_str ?? "<null>"}");

  // ...
}

```
