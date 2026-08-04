using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several checkbox content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Write a label before the checkbox.
            builder.Writeln($"Item {i}:");

            // Create an inline checkbox StructuredDocumentTag.
            StructuredDocumentTag checkBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
            {
                Title = $"Option{i}",
                Tag = $"opt{i}",
                Checked = i % 2 == 0 // Even items are checked.
            };

            // Insert the checkbox into the document.
            builder.InsertNode(checkBox);
            builder.Writeln(); // Move to the next line.
        }

        // Save the sample document.
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // Load the document (simulating processing an existing file).
        Document loadedDoc = new Document(samplePath);

        // Find all checkbox content controls.
        List<StructuredDocumentTag> checkBoxTags = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(tag => tag.SdtType == SdtType.Checkbox)
            .ToList();

        // Prepare CSV lines.
        List<string> csvLines = new List<string>
        {
            "Title,Tag,Checked"
        };

        foreach (StructuredDocumentTag tag in checkBoxTags)
        {
            string title = tag.Title ?? string.Empty;
            string tagValue = tag.Tag ?? string.Empty;
            string checkedState = tag.Checked.ToString();
            csvLines.Add($"{EscapeCsv(title)},{EscapeCsv(tagValue)},{checkedState}");
        }

        // Write the CSV file.
        const string csvPath = "checkboxes.csv";
        File.WriteAllLines(csvPath, csvLines);
    }

    // Simple CSV escaping for values that may contain commas or quotes.
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
