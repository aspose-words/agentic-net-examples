using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several checkbox content controls.
        Document doc = new Document();
        // First checkbox
        Paragraph para1 = new Paragraph(doc);
        StructuredDocumentTag checkbox1 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkbox1.Title = "AcceptTerms";
        checkbox1.Tag = "accept-terms";
        checkbox1.Checked = true;
        para1.AppendChild(checkbox1);
        doc.FirstSection.Body.AppendChild(para1);
        // Second checkbox
        Paragraph para2 = new Paragraph(doc);
        StructuredDocumentTag checkbox2 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkbox2.Title = "SubscribeNewsletter";
        checkbox2.Tag = "subscribe-newsletter";
        checkbox2.Checked = false;
        para2.AppendChild(checkbox2);
        doc.FirstSection.Body.AppendChild(para2);
        // Third checkbox
        Paragraph para3 = new Paragraph(doc);
        StructuredDocumentTag checkbox3 = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
        checkbox3.Title = "EnableNotifications";
        checkbox3.Tag = "enable-notifications";
        checkbox3.Checked = true;
        para3.AppendChild(checkbox3);
        doc.FirstSection.Body.AppendChild(para3);
        // Save the sample document (optional, demonstrates persistence).
        const string samplePath = "sample.docx";
        doc.Save(samplePath);
        // Load the document (could reuse the same instance, but loading demonstrates typical workflow).
        Document loadedDoc = new Document(samplePath);
        // Find all checkbox content controls.
        var checkboxes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Checkbox)
            .ToList();
        // Prepare CSV content.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Title,Tag,Checked");
        foreach (var cb in checkboxes)
        {
            string title = EscapeForCsv(cb.Title ?? string.Empty);
            string tag = EscapeForCsv(cb.Tag ?? string.Empty);
            string checkedValue = cb.Checked.ToString();
            csvBuilder.AppendLine($"{title},{tag},{checkedValue}");
        }
        // Write CSV to file.
        const string csvPath = "checkboxes.csv";
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }
    // Helper to escape CSV fields containing commas or quotes.
    private static string EscapeForCsv(string field)
    {
        if (field.Contains('"'))
        {
            field = field.Replace("\"", "\"\"");
        }
        if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
        {
            field = $"\"{field}\"";
        }
        return field;
    }
}
