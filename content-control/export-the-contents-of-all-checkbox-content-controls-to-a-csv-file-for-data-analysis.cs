using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json; // Included as required package

public class Program
{
    public static void Main()
    {
        const string inputPath = "checkboxes.docx";
        const string csvPath = "checkboxes.csv";

        // Create a sample document with checkbox content controls if it does not exist.
        if (!File.Exists(inputPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            AddCheckbox(doc, builder, "AgreeTerms", "agree-terms", true);
            AddCheckbox(doc, builder, "SubscribeNewsletter", "subscribe-newsletter", false);
            AddCheckbox(doc, builder, "ReceiveUpdates", "receive-updates", true);

            doc.Save(inputPath);
        }

        // Load the document that contains the checkboxes.
        Document loadedDoc = new Document(inputPath);

        // Retrieve all checkbox StructuredDocumentTag nodes.
        var checkboxes = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Checkbox)
            .ToList();

        // Build CSV content.
        var sb = new StringBuilder();
        sb.AppendLine("Title,Tag,Checked");
        foreach (var cb in checkboxes)
        {
            string title = cb.Title ?? string.Empty;
            string tag = cb.Tag ?? string.Empty;
            string checkedValue = cb.Checked ? "True" : "False";

            sb.AppendLine($"{EscapeCsv(title)},{EscapeCsv(tag)},{checkedValue}");
        }

        // Write CSV file.
        File.WriteAllText(csvPath, sb.ToString());

        // Inform about the result (no interactive input required).
        Console.WriteLine($"Exported {checkboxes.Count} checkbox(es) to \"{csvPath}\".");
    }

    // Helper to insert a checkbox content control with title, tag and state.
    private static void AddCheckbox(Document doc, DocumentBuilder builder, string title, string tag, bool isChecked)
    {
        // Write a label before the checkbox.
        builder.Write($"{title}: ");

        // Create the checkbox StructuredDocumentTag.
        StructuredDocumentTag checkbox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = title,
            Tag = tag,
            Checked = isChecked
        };

        // Insert the checkbox at the current builder position.
        builder.InsertNode(checkbox);

        // Move to a new line for the next control.
        builder.Writeln();
    }

    // Simple CSV field escaper.
    private static string EscapeCsv(string field)
    {
        if (field.Contains("\""))
            field = field.Replace("\"", "\"\"");

        if (field.Contains(",") || field.Contains("\"") || field.Contains("\r") || field.Contains("\n"))
            return $"\"{field}\"";

        return field;
    }
}
