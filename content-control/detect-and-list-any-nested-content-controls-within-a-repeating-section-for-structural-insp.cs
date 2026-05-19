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
        // -----------------------------------------------------------------
        // 1. Create a sample document with a repeating section that contains
        //    nested plain‑text and rich‑text content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the repeating section (block level) and add it to the body.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block);
        repeatingSection.Title = "RepeatingSection";
        repeatingSection.Tag = "repeating-section";

        // The repeating section must contain at least one block element (a paragraph).
        Paragraph innerParagraph = new Paragraph(doc);
        repeatingSection.AppendChild(innerParagraph);
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // Move the builder to the paragraph that lives inside the repeating section.
        builder.MoveTo(innerParagraph);

        // ----- Nested plain‑text content control (inline) -----
        StructuredDocumentTag nestedPlain = builder.InsertStructuredDocumentTag(SdtType.PlainText);
        nestedPlain.Title = "NestedPlain";
        nestedPlain.Tag = "nested-plain";
        // Write the text that will be inside the plain‑text SDT.
        builder.Write("Plain inside repeating");

        // Add a space between the two controls for readability.
        builder.Write(" ");

        // ----- Nested rich‑text content control (inline) -----
        StructuredDocumentTag nestedRich = builder.InsertStructuredDocumentTag(SdtType.RichText);
        nestedRich.Title = "NestedRich";
        nestedRich.Tag = "nested-rich";
        // Write the text that will be inside the rich‑text SDT.
        builder.Write("Rich inside repeating");

        // -----------------------------------------------------------------
        // 2. Save the sample document.
        // -----------------------------------------------------------------
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and enumerate nested content controls inside
        //    each repeating section.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Find all repeating‑section SDTs.
        var repeatingControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.RepeatingSection)
            .ToList();

        var reportItems = new List<NestedControlInfo>();

        foreach (var repeating in repeatingControls)
        {
            // All descendant SDTs inside the current repeating section.
            var nestedControls = repeating
                .GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .ToList();

            foreach (var nested in nestedControls)
            {
                // Compute depth relative to the repeating section.
                int depth = 0;
                Node parent = nested.ParentNode;
                while (parent != null && parent != repeating)
                {
                    depth++;
                    parent = parent.ParentNode;
                }

                reportItems.Add(new NestedControlInfo
                {
                    RepeatingSectionTitle = repeating.Title,
                    RepeatingSectionTag = repeating.Tag,
                    NestedTitle = nested.Title,
                    NestedTag = nested.Tag,
                    NestedType = nested.SdtType.ToString(),
                    DepthFromRepeating = depth
                });
            }
        }

        // -----------------------------------------------------------------
        // 4. Serialize the report to JSON and write it to a file.
        // -----------------------------------------------------------------
        string jsonReport = JsonConvert.SerializeObject(reportItems, Formatting.Indented);
        const string jsonPath = "nested_sdt_report.json";
        File.WriteAllText(jsonPath, jsonReport);

        // Output the report to the console.
        Console.WriteLine("Nested Content Controls within Repeating Sections:");
        Console.WriteLine(jsonReport);
    }

    // Helper class for JSON serialization.
    private class NestedControlInfo
    {
        public string RepeatingSectionTitle { get; set; } = string.Empty;
        public string RepeatingSectionTag { get; set; } = string.Empty;
        public string NestedTitle { get; set; } = string.Empty;
        public string NestedTag { get; set; } = string.Empty;
        public string NestedType { get; set; } = string.Empty;
        public int DepthFromRepeating { get; set; }
    }
}
