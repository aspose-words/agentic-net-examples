using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Paths for the sample document and the CSV output.
        const string docPath = "sample.docx";
        const string csvPath = "checkboxes.csv";

        // -------------------------------------------------
        // Create a sample DOCX containing several checkbox content controls.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Inline checkbox 1.
        StructuredDocumentTag checkBox1 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeTerms",
            Tag = "agree",
            Checked = true
        };
        builder.InsertNode(checkBox1);
        builder.Writeln();

        // Inline checkbox 2.
        StructuredDocumentTag checkBox2 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "SubscribeNewsletter",
            Tag = "subscribe",
            Checked = false
        };
        builder.InsertNode(checkBox2);
        builder.Writeln();

        // Block-level checkbox.
        StructuredDocumentTag checkBox3 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Block)
        {
            Title = "AcceptPolicy",
            Tag = "policy",
            Checked = true
        };
        // Block-level SDTs must contain at least one paragraph.
        Paragraph placeholderParagraph = new Paragraph(doc);
        placeholderParagraph.AppendChild(new Run(doc, " "));
        checkBox3.AppendChild(placeholderParagraph);
        doc.FirstSection.Body.AppendChild(checkBox3);

        // Save the sample document.
        doc.Save(docPath);

        // -------------------------------------------------
        // Load the document and export checkbox data to CSV.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);

        var checkboxData = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Checkbox)
            .Select(sdt => new
            {
                Title = sdt.Title ?? string.Empty,
                Tag = sdt.Tag ?? string.Empty,
                Checked = sdt.Checked
            })
            .ToList();

        using (StreamWriter writer = new StreamWriter(csvPath, false))
        {
            writer.WriteLine("Title,Tag,Checked");
            foreach (var item in checkboxData)
            {
                writer.WriteLine($"{EscapeCsv(item.Title)},{EscapeCsv(item.Tag)},{item.Checked}");
            }
        }

        // Inform the user (no interactive input required).
        Console.WriteLine($"Exported {checkboxData.Count} checkbox(es) to \"{csvPath}\".");
    }

    // Simple CSV field escaper.
    private static string EscapeCsv(string value)
    {
        if (value.Contains(',') || value.Contains('\"') || value.Contains('\n') || value.Contains('\r'))
        {
            string escaped = value.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return value;
    }
}
