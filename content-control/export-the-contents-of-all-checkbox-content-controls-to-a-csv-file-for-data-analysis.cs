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
        Document sampleDoc = new Document();
        Paragraph para = sampleDoc.FirstSection.Body.FirstParagraph;

        // First checkbox (checked) – inline.
        StructuredDocumentTag checkbox1 = new StructuredDocumentTag(sampleDoc, SdtType.Checkbox, MarkupLevel.Inline);
        checkbox1.Title = "AcceptTerms";
        checkbox1.Tag = "accept_terms";
        checkbox1.Checked = true;
        para.AppendChild(checkbox1);
        para.AppendChild(new Run(sampleDoc, " ")); // space separator

        // Second checkbox (unchecked) – inline.
        StructuredDocumentTag checkbox2 = new StructuredDocumentTag(sampleDoc, SdtType.Checkbox, MarkupLevel.Inline);
        checkbox2.Title = "SubscribeNewsletter";
        checkbox2.Tag = "subscribe_newsletter";
        checkbox2.Checked = false;
        para.AppendChild(checkbox2);
        para.AppendChild(new Run(sampleDoc, " "));

        // Third checkbox (checked) – block level.
        StructuredDocumentTag checkbox3 = new StructuredDocumentTag(sampleDoc, SdtType.Checkbox, MarkupLevel.Block);
        checkbox3.Title = "ReceiveUpdates";
        checkbox3.Tag = "receive_updates";
        checkbox3.Checked = true;

        // Block‑level SDT must contain at least one paragraph.
        Paragraph innerPara = new Paragraph(sampleDoc);
        innerPara.AppendChild(new Run(sampleDoc, " "));
        checkbox3.AppendChild(innerPara);

        // Append the block‑level SDT directly to the document body (not inside a paragraph).
        sampleDoc.FirstSection.Body.AppendChild(checkbox3);

        // Save the sample document.
        const string samplePath = "sample.docx";
        sampleDoc.Save(samplePath);

        // Load the document for processing.
        Document doc = new Document(samplePath);

        // Find all checkbox content controls.
        var checkboxes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Checkbox)
            .Select(sdt => new
            {
                Title = sdt.Title ?? string.Empty,
                Tag = sdt.Tag ?? string.Empty,
                Checked = sdt.Checked
            })
            .ToList();

        // Prepare CSV content.
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Title,Tag,Checked");

        // Helper to escape CSV fields.
        string Escape(string field)
        {
            if (field.Contains('"') || field.Contains(','))
                return $"\"{field.Replace("\"", "\"\"")}\"";
            return field;
        }

        foreach (var cb in checkboxes)
        {
            string line = $"{Escape(cb.Title)},{Escape(cb.Tag)},{cb.Checked}";
            csvBuilder.AppendLine(line);
        }

        // Write CSV to file.
        const string csvPath = "checkboxes.csv";
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }
}
