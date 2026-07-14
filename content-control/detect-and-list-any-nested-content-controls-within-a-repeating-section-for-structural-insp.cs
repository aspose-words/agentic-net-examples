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
        // Create a new document.
        Document doc = new Document();
        // Create a repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "RepeatingSection",
            Tag = "repeating-section"
        };

        // Add a header paragraph inside the repeating section.
        Paragraph headerParagraph = new Paragraph(doc);
        headerParagraph.AppendChild(new Run(doc, "Repeating Section Header"));
        repeatingSection.AppendChild(headerParagraph);

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

        // Add a paragraph with an inline plain text content control.
        Paragraph inlineParagraph = new Paragraph(doc);
        StructuredDocumentTag nestedInline = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NestedInline",
            Tag = "nested-inline"
        };
        nestedInline.AppendChild(new Run(doc, "Inline nested content"));
        inlineParagraph.AppendChild(nestedInline);
        repeatingSection.AppendChild(inlineParagraph);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Analyze the document for nested content controls within repeating sections.
        List<NestedControlInfo> results = new List<NestedControlInfo>();

        IEnumerable<StructuredDocumentTag> allSdt = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>();

        foreach (StructuredDocumentTag repeating in allSdt.Where(s => s.SdtType == SdtType.RepeatingSection))
        {
            // Find nested SDTs inside this repeating section (excluding the repeating section itself).
            IEnumerable<StructuredDocumentTag> nestedSdts = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(s => s != repeating);

            foreach (StructuredDocumentTag nested in nestedSdts)
            {
                results.Add(new NestedControlInfo
                {
                    RepeatingSectionTitle = repeating.Title,
                    RepeatingSectionTag = repeating.Tag,
                    NestedTitle = nested.Title,
                    NestedTag = nested.Tag,
                    NestedType = nested.SdtType.ToString(),
                    NestedText = nested.GetText().Trim()
                });
            }
        }

        // Serialize the inspection results to JSON.
        string json = JsonConvert.SerializeObject(results, Formatting.Indented);
        const string jsonPath = "nested-content-controls.json";
        File.WriteAllText(jsonPath, json);

        // Optionally, write a short summary to the console.
        Console.WriteLine($"Document saved to '{docPath}'.");
        Console.WriteLine($"Inspection results saved to '{jsonPath}'.");
        Console.WriteLine($"Found {results.Count} nested content control(s) inside repeating sections.");
    }

    // Helper class to hold inspection data.
    private class NestedControlInfo
    {
        public string RepeatingSectionTitle { get; set; } = string.Empty;
        public string RepeatingSectionTag { get; set; } = string.Empty;
        public string NestedTitle { get; set; } = string.Empty;
        public string NestedTag { get; set; } = string.Empty;
        public string NestedType { get; set; } = string.Empty;
        public string NestedText { get; set; } = string.Empty;
    }
}
