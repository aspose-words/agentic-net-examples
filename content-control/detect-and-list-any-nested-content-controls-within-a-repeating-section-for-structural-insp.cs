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
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
        {
            Title = "RepeatingSection",
            Tag = "rep-section"
        };

        // Create a nested block-level rich text content control.
        StructuredDocumentTag nestedBlock = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "NestedBlock",
            Tag = "nested-block"
        };
        Paragraph blockParagraph = new Paragraph(doc);
        blockParagraph.AppendChild(new Run(doc, "Content inside nested block-level SDT."));
        nestedBlock.AppendChild(blockParagraph);

        // Create a nested inline plain text content control.
        StructuredDocumentTag nestedInline = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "NestedInline",
            Tag = "nested-inline"
        };
        nestedInline.RemoveAllChildren();
        nestedInline.AppendChild(new Run(doc, "Inline SDT text."));

        // Add the inline SDT inside a paragraph.
        Paragraph inlineParagraph = new Paragraph(doc);
        inlineParagraph.AppendChild(nestedInline);

        // Assemble the repeating section: add both nested controls.
        repeatingSection.AppendChild(nestedBlock);
        repeatingSection.AppendChild(inlineParagraph);

        // Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Save the sample document.
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Detect nested content controls within each repeating section.
        List<NestedSdtInfo> report = new List<NestedSdtInfo>();

        // Find all repeating section SDTs in the document.
        IEnumerable<StructuredDocumentTag> repeatingSdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

        foreach (StructuredDocumentTag repeating in repeatingSdtNodes)
        {
            // Find all descendant SDTs inside the repeating section.
            IEnumerable<StructuredDocumentTag> nestedSdts = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>();

            foreach (StructuredDocumentTag nested in nestedSdts)
            {
                report.Add(new NestedSdtInfo
                {
                    RepeatingSectionTitle = repeating.Title,
                    RepeatingSectionTag = repeating.Tag,
                    NestedTitle = nested.Title,
                    NestedTag = nested.Tag,
                    NestedType = nested.SdtType.ToString()
                });
            }
        }

        // Serialize the report to JSON.
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string jsonPath = "nested-content-controls.json";
        File.WriteAllText(jsonPath, json);

        // Output result to console.
        Console.WriteLine("Nested content controls detection completed.");
        Console.WriteLine($"Document saved as: {Path.GetFullPath(docPath)}");
        Console.WriteLine($"Report saved as: {Path.GetFullPath(jsonPath)}");
        Console.WriteLine("Report content:");
        Console.WriteLine(json);
    }

    // Helper class to hold information about nested SDTs.
    private class NestedSdtInfo
    {
        public string RepeatingSectionTitle { get; set; } = string.Empty;
        public string RepeatingSectionTag { get; set; } = string.Empty;
        public string NestedTitle { get; set; } = string.Empty;
        public string NestedTag { get; set; } = string.Empty;
        public string NestedType { get; set; } = string.Empty;
    }
}
