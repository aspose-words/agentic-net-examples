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
        // Create a new blank document.
        Document doc = new Document();

        // Create a repeating section content control (block level).
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "RepeatingSection";
        repeatingSection.Tag = "repeating-section";

        // First paragraph inside the repeating section.
        Paragraph paragraph1 = new Paragraph(doc);
        paragraph1.AppendChild(new Run(doc, "Item 1: "));
        // Nested plain text content control inside the first paragraph.
        StructuredDocumentTag nestedPlain = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
        nestedPlain.Title = "NestedPlain";
        nestedPlain.Tag = "nested-plain";
        nestedPlain.RemoveAllChildren();
        nestedPlain.AppendChild(new Run(doc, "ValueA"));
        paragraph1.AppendChild(nestedPlain);
        repeatingSection.AppendChild(paragraph1);

        // Second paragraph inside the repeating section.
        Paragraph paragraph2 = new Paragraph(doc);
        paragraph2.AppendChild(new Run(doc, "Item 2: "));
        // Nested rich text content control inside the second paragraph.
        StructuredDocumentTag nestedRich = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Inline);
        nestedRich.Title = "NestedRich";
        nestedRich.Tag = "nested-rich";
        nestedRich.RemoveAllChildren();
        nestedRich.AppendChild(new Run(doc, "ValueB"));
        paragraph2.AppendChild(nestedRich);
        repeatingSection.AppendChild(paragraph2);

        // Add the repeating section to the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string docPath = "RepeatingSection.docx";
        doc.Save(docPath);

        // Detect nested content controls within each repeating section.
        List<NestedControlInfo> nestedControls = new List<NestedControlInfo>();

        IEnumerable<StructuredDocumentTag> repeatingSections = doc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

        foreach (StructuredDocumentTag repeating in repeatingSections)
        {
            // Get all descendant StructuredDocumentTag nodes (nested controls).
            IEnumerable<StructuredDocumentTag> descendants = repeating
                .GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>();

            foreach (StructuredDocumentTag nested in descendants)
            {
                var info = new NestedControlInfo
                {
                    ParentRepeatingTitle = repeating.Title,
                    ParentRepeatingTag = repeating.Tag,
                    Title = nested.Title,
                    Tag = nested.Tag,
                    SdtType = nested.SdtType.ToString(),
                    Text = nested.GetText().Trim()
                };
                nestedControls.Add(info);
                Console.WriteLine($"Nested SDT - Title: {info.Title}, Tag: {info.Tag}, Type: {info.SdtType}, Text: \"{info.Text}\"");
            }
        }

        // Serialize the detection result to JSON.
        string json = JsonConvert.SerializeObject(nestedControls, Formatting.Indented);
        const string jsonPath = "NestedControls.json";
        File.WriteAllText(jsonPath, json);
    }

    // Helper class for JSON serialization.
    private class NestedControlInfo
    {
        public string ParentRepeatingTitle { get; set; } = string.Empty;
        public string ParentRepeatingTag { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Tag { get; set; } = string.Empty;
        public string SdtType { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
    }
}
