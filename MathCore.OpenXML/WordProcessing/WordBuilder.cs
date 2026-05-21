using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using MathCore.OpenXML.WordProcessing.Extensions.Word;

namespace MathCore.OpenXML.WordProcessing;

/// <summary>
/// Fluent-конструктор нового Word-документа.
/// </summary>
/// <example>
/// <code>
/// Word.Create()
///    .Paragraph("Отчет")
///    .Paragraph("Сформировано автоматически", p =&gt; p.AlignRight().Italic())
///    .Table(
///        new[] { "ID", "Name" },
///        new[] { "1", "Item" })
///    .SaveTo("report.docx");
/// </code>
/// </example>
public class WordBuilder
{
    private readonly Body _Body = new();

    /// <summary>Добавить абзац с указанным текстом.</summary>
    public WordBuilder Paragraph(string Text)
    {
        _Body.AppendChild(new Paragraph().Add(Text));
        return this;
    }

    /// <summary>Добавить абзац и настроить его через callback.</summary>
    public WordBuilder Paragraph(string Text, Action<Paragraph> Configure)
    {
        var paragraph = new Paragraph().Add(Text);
        Configure(paragraph);
        _Body.AppendChild(paragraph);

        return this;
    }

    /// <summary>Добавить абзац, полностью формируемый в callback.</summary>
    public WordBuilder Paragraph(Action<Paragraph> Build)
    {
        var paragraph = new Paragraph();
        Build(paragraph);
        _Body.AppendChild(paragraph);

        return this;
    }

    /// <summary>Добавить простую таблицу из набора строк и ячеек.</summary>
    public WordBuilder Table(IEnumerable<IEnumerable<string?>> Rows)
    {
        var table = new Table();
        table.Add(new TableProperties(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4 },
            new BottomBorder { Val = BorderValues.Single, Size = 4 },
            new LeftBorder { Val = BorderValues.Single, Size = 4 },
            new RightBorder { Val = BorderValues.Single, Size = 4 },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        foreach (var row_cells in Rows)
        {
            var row = new TableRow();
            foreach (var cell_text in row_cells)
            {
                var paragraph = new Paragraph().Add(cell_text ?? string.Empty);
                row.AppendChild(new TableCell(paragraph));
            }

            table.Add(row);
        }

        _Body.AppendChild(table);
        return this;
    }

    /// <summary>Добавить простую таблицу из массива строк.</summary>
    public WordBuilder Table(params string?[][] Rows) => Table((IEnumerable<IEnumerable<string?>>)Rows);

    /// <summary>Добавить таблицу и настроить ее вручную через callback.</summary>
    public WordBuilder Table(Action<Table> Build)
    {
        var table = new Table();
        Build(table);

        _Body.AppendChild(table);
        return this;
    }

    /// <summary>Добавить свойства секции документа.</summary>
    public WordBuilder Section(Action<SectionProperties> Build)
    {
        var section = new SectionProperties();
        Build(section);

        _Body.AppendChild(section);
        return this;
    }

    /// <summary>Сохранить созданный документ в файл по указанному пути.</summary>
    public FileInfo SaveTo(string FilePath) => SaveTo(new FileInfo(FilePath));

    /// <summary>Сохранить созданный документ в указанный файл.</summary>
    public FileInfo SaveTo(FileInfo File)
    {
        using var document = WordprocessingDocument.Create(File.FullName, WordprocessingDocumentType.Document);
        SaveTo(document);

        File.Refresh();
        return File;
    }

    /// <summary>Сохранить созданный документ в поток.</summary>
    public void SaveTo(Stream Stream)
    {
        using var document = WordprocessingDocument.Create(Stream, WordprocessingDocumentType.Document);
        SaveTo(document);
    }

    private void SaveTo(WordprocessingDocument Document)
    {
        var main_part = Document.AddMainDocumentPart();
        main_part.Document = new() { Body = (Body)_Body.CloneNode(true) };
    }
}
