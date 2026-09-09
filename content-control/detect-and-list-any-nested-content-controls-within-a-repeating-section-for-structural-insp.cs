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

        // Build a repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "RepeatingSection",
            Tag = "rep-section"
        };

        // Add a simple paragraph inside the repeating section.
        Paragraph startParagraph = new Paragraph(doc);
        startParagraph.AppendChild(new Run(doc, "Repeating item start"));
        repeatingSection.AppendChild(startParagraph);

        // Add a nested block-level rich text content control.
        StructuredDocumentTag nestedBlock = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "NestedBlock",
            Tag = "nested-block"
        };
        Paragraph blockParagraph = new Paragraph(doc);
        blockParagraph.AppendChild(new Run(doc, "Nested block content"));
        nestedBlock.AppendChild(blockParagraph);
        repeatingSection.AppendChild(nestedBlock);

        // Add a nested inline plain text content control inside a paragraph.
        Paragraph inlineParagraph = new Paragraph(doc);
        StructuredDocumentTag nestedInline = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NestedInline",
            Tag = "nested-inline"
        };
        nestedInline.RemoveAllChildren();
        nestedInline.AppendChild(new Run(doc, "Inline content"));
        inlineParagraph.AppendChild(nestedInline);
        repeatingSection.AppendChild(inlineParagraph);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string docPath = "NestedRepeatingSection.docx";
        doc.Save(docPath);

        // Detect nested content controls within each repeating section.
        var report = new List<NestedControlInfo>();

        IEnumerable<StructuredDocumentTag> repeatingControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

        foreach (StructuredDocumentTag repeating in repeatingControls)
        {
            // Find all descendant StructuredDocumentTag nodes that are not the repeating section itself.
            IEnumerable<StructuredDocumentTag> nestedControls = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(sdt => sdt != repeating);

            foreach (StructuredDocumentTag nested in nestedControls)
            {
                report.Add(new NestedControlInfo
                {
                    ParentRepeatingTitle = repeating.Title,
                    ParentRepeatingTag = repeating.Tag,
                    NestedTitle = nested.Title,
                    NestedTag = nested.Tag,
                    NestedType = nested.SdtType.ToString()
                });
            }
        }

        // Serialize the inspection result to JSON.
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string jsonPath = "NestedControls.json";
        File.WriteAllText(jsonPath, json);

        // Output the result to the console.
        Console.WriteLine("Nested content controls detected within repeating sections:");
        Console.WriteLine(json);
    }

    // Helper class for JSON serialization.
    private class NestedControlInfo
    {
        public string ParentRepeatingTitle { get; set; } = string.Empty;
        public string ParentRepeatingTag { get; set; } = string.Empty;
        public string NestedTitle { get; set; } = string.Empty;
        public string NestedTag { get; set; } = string.Empty;
        public string NestedType { get; set; } = string.Empty;
    }
}
