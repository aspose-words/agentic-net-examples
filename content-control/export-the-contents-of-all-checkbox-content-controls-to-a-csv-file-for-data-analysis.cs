using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several checkbox content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First checkbox (checked)
        StructuredDocumentTag checkBox1 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Checked = true,
            Title = "Accept Terms",
            Tag = "AcceptTerms"
        };
        builder.InsertNode(checkBox1);
        builder.Writeln(); // Move to next line

        // Second checkbox (unchecked)
        StructuredDocumentTag checkBox2 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Checked = false,
            Title = "Subscribe Newsletter",
            Tag = "Subscribe"
        };
        builder.InsertNode(checkBox2);
        builder.Writeln();

        // Third checkbox (checked)
        StructuredDocumentTag checkBox3 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Checked = true,
            Title = "Enable Notifications",
            Tag = "EnableNotif"
        };
        builder.InsertNode(checkBox3);
        builder.Writeln();

        // Save the sample document (optional, demonstrates lifecycle usage).
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // Load the document (could reuse the same instance, but follows load rule).
        Document loadedDoc = new Document(samplePath);

        // Prepare CSV output.
        const string csvPath = "checkboxes.csv";
        using (StreamWriter writer = new StreamWriter(csvPath, false, Encoding.UTF8))
        {
            // Write CSV header.
            writer.WriteLine("Title,Tag,Checked");

            // Enumerate all StructuredDocumentTag nodes.
            NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            foreach (Node node in sdtNodes)
            {
                if (node is StructuredDocumentTag sdt && sdt.SdtType == SdtType.Checkbox)
                {
                    string title = sdt.Title ?? string.Empty;
                    string tag = sdt.Tag ?? string.Empty;
                    string checkedValue = sdt.Checked ? "True" : "False";

                    // Simple CSV escaping: wrap each field in quotes and double any internal quotes.
                    writer.WriteLine($"{EscapeCsv(title)},{EscapeCsv(tag)},{checkedValue}");
                }
            }
        }

        // The program finishes here; no interactive prompts are used.
    }

    // Helper method to escape CSV fields.
    private static string EscapeCsv(string field)
    {
        if (field.Contains("\"") || field.Contains(",") || field.Contains("\n") || field.Contains("\r"))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }
        return field;
    }
}
