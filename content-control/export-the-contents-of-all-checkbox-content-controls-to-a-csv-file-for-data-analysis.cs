using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX that contains a few checkbox content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First checkbox (checked)
        builder.Writeln("Task 1:");
        StructuredDocumentTag checkBox1 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Task1",
            Tag = "task1",
            Checked = true
        };
        builder.InsertNode(checkBox1);
        builder.Writeln(" Completed");

        // Second checkbox (unchecked)
        builder.Writeln();
        builder.Writeln("Task 2:");
        StructuredDocumentTag checkBox2 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Task2",
            Tag = "task2",
            Checked = false
        };
        builder.InsertNode(checkBox2);
        builder.Writeln(" Pending");

        // Save the sample document.
        const string samplePath = "SampleCheckboxes.docx";
        doc.Save(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the document and extract all checkbox content controls.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);

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

        // -----------------------------------------------------------------
        // 3. Write the extracted data to a CSV file.
        // -----------------------------------------------------------------
        const string csvPath = "CheckboxExport.csv";
        var csvLines = new List<string> { "Title,Tag,Checked" };
        foreach (var item in checkboxData)
        {
            csvLines.Add($"{EscapeCsv(item.Title)},{EscapeCsv(item.Tag)},{item.Checked}");
        }

        File.WriteAllLines(csvPath, csvLines);
    }

    // Simple CSV escaping for values that may contain commas, quotes or newlines.
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
