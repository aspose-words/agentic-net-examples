using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a plain‑text content control.
        StructuredDocumentTag plainText = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainTextControl",
            Tag = "plain"
        };
        plainText.RemoveAllChildren();
        plainText.AppendChild(new Run(doc, "Plain text value"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(plainText);

        // Insert a rich‑text content control.
        StructuredDocumentTag richText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "RichTextControl",
            Tag = "rich"
        };
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Rich text value"));
        richText.AppendChild(richParagraph);
        doc.FirstSection.Body.AppendChild(richText);

        // Insert a checkbox content control.
        StructuredDocumentTag checkBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "CheckBoxControl",
            Tag = "checkbox",
            Checked = true
        };
        doc.FirstSection.Body.FirstParagraph.AppendChild(checkBox);

        // Insert a drop‑down list content control.
        StructuredDocumentTag dropDown = new StructuredDocumentTag(doc, SdtType.DropDownList, MarkupLevel.Inline)
        {
            Title = "DropDownControl",
            Tag = "dropdown"
        };
        dropDown.ListItems.Add(new SdtListItem("Option 1", "1"));
        dropDown.ListItems.Add(new SdtListItem("Option 2", "2"));
        doc.FirstSection.Body.FirstParagraph.AppendChild(dropDown);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Load the document back (demonstrates load operation).
        Document loadedDoc = new Document(docPath);

        // Collect information about each content control.
        var reportItems = new List<ContentControlInfo>();
        NodeCollection sdtNodes = loadedDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        foreach (Node node in sdtNodes)
        {
            if (node is StructuredDocumentTag sdt)
            {
                var info = new ContentControlInfo
                {
                    Title = sdt.Title,
                    Tag = sdt.Tag,
                    Type = sdt.SdtType.ToString(),
                    Appearance = sdt.Appearance.ToString(),
                    IsShowingPlaceholderText = sdt.IsShowingPlaceholderText
                };
                reportItems.Add(info);
            }
        }

        // Serialize the report to JSON.
        string jsonReport = JsonConvert.SerializeObject(reportItems, Formatting.Indented);
        const string jsonPath = "content-controls-report.json";
        File.WriteAllText(jsonPath, jsonReport);

        // Output the report to the console.
        Console.WriteLine("Content Control Summary:");
        Console.WriteLine(jsonReport);
    }

    // Simple DTO for JSON serialization.
    private class ContentControlInfo
    {
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Appearance { get; set; } = string.Empty;
        public bool IsShowingPlaceholderText { get; set; }
    }
}
