using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for inserting content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first checkbox content control.
        builder.Writeln("Option 1:");
        StructuredDocumentTag checkBox1 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Option1",
            Tag = "opt1",
            Checked = true
        };
        builder.InsertNode(checkBox1);
        builder.Writeln();

        // Insert second checkbox content control.
        builder.Writeln("Option 2:");
        StructuredDocumentTag checkBox2 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Option2",
            Tag = "opt2",
            Checked = false
        };
        builder.InsertNode(checkBox2);
        builder.Writeln();

        // Insert third checkbox content control.
        builder.Writeln("Option 3:");
        StructuredDocumentTag checkBox3 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Option3",
            Tag = "opt3",
            Checked = true
        };
        builder.InsertNode(checkBox3);
        builder.Writeln();

        // Save the sample document (optional, demonstrates persistence).
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Enumerate all checkbox StructuredDocumentTag nodes in the document.
        var checkboxTags = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                              .OfType<StructuredDocumentTag>()
                              .Where(sdt => sdt.SdtType == SdtType.Checkbox)
                              .ToList();

        // Prepare CSV content.
        var csvLines = new[]
        {
            "Title,Tag,Checked"
        }
        .Concat(checkboxTags.Select(sdt =>
            $"{EscapeCsv(sdt.Title ?? string.Empty)},{EscapeCsv(sdt.Tag ?? string.Empty)},{sdt.Checked}"));

        // Write CSV to file.
        const string csvPath = "checkboxes.csv";
        File.WriteAllLines(csvPath, csvLines);
    }

    // Helper to escape CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(',') || field.Contains('\"') || field.Contains('\n') || field.Contains('\r'))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
