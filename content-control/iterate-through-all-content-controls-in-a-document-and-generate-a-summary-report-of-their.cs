using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document with various content controls.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Plain text inline content control.
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "John Doe"));
        builder.InsertNode(plainTextSdt);

        // Rich text block content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Comments",
            Tag = "comments"
        };
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "This is a comment."));
        richTextSdt.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Checkbox inline content control.
        StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "Agree",
            Tag = "agree",
            Checked = true
        };
        builder.InsertNode(checkboxSdt);

        // Drop‑down list inline content control.
        StructuredDocumentTag dropdownSdt = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "Options",
            Tag = "options"
        };
        dropdownSdt.ListItems.Add(new SdtListItem("Option A", "A"));
        dropdownSdt.ListItems.Add(new SdtListItem("Option B", "B"));
        builder.InsertNode(dropdownSdt);

        // Date inline content control.
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "Date",
            Tag = "date",
            DateDisplayFormat = "yyyy-MM-dd"
        };
        builder.InsertNode(dateSdt);

        // Save the sample document.
        const string samplePath = "sample.docx";
        doc.Save(samplePath);

        // Step 2: Load the document and enumerate all content controls.
        Document loadedDoc = new Document(samplePath);
        IEnumerable<StructuredDocumentTag> sdtNodes = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>();

        // Prepare a simple DTO for JSON serialization.
        var reportItems = new List<ControlInfo>();
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            reportItems.Add(new ControlInfo
            {
                Id = sdt.Id,
                Type = sdt.SdtType.ToString(),
                Title = sdt.Title,
                Tag = sdt.Tag
            });
        }

        // Serialize the report to JSON.
        string jsonReport = JsonConvert.SerializeObject(reportItems, Formatting.Indented);
        const string reportPath = "content_controls_report.json";
        File.WriteAllText(reportPath, jsonReport);

        // Output a brief summary to the console.
        Console.WriteLine($"Found {reportItems.Count} content controls. Report saved to '{reportPath}'.");
    }

    // DTO used for the JSON report.
    private class ControlInfo
    {
        public int Id { get; set; }
        public string Type { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
    }
}
