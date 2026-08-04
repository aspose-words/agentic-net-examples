using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace ContentControlInspection
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a repeating section that contains nested content controls.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create the outer repeating section (block level).
            StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
            {
                Title = "RepeatingSection",
                Tag = "rep-section"
            };
            // Add a placeholder paragraph inside the repeating section.
            Paragraph placeholderParagraph = new Paragraph(doc);
            placeholderParagraph.AppendChild(new Run(doc, "Repeating section placeholder"));
            repeatingSection.AppendChild(placeholderParagraph);
            doc.FirstSection.Body.AppendChild(repeatingSection);

            // Create a repeating section item.
            StructuredDocumentTag repeatingItem = new StructuredDocumentTag(doc, SdtType.RepeatingSectionItem, MarkupLevel.Block);
            repeatingSection.AppendChild(repeatingItem);

            // Inside the item, add a nested plain‑text content control (inline level).
            StructuredDocumentTag nestedPlain = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "NestedPlain",
                Tag = "nested-plain"
            };
            nestedPlain.RemoveAllChildren();
            nestedPlain.AppendChild(new Run(doc, "Nested value"));

            // The inline SDT must be placed inside a paragraph.
            Paragraph innerParagraph = new Paragraph(doc);
            innerParagraph.AppendChild(new Run(doc, "Before nested "));
            innerParagraph.AppendChild(nestedPlain);
            innerParagraph.AppendChild(new Run(doc, " after nested."));
            repeatingItem.AppendChild(innerParagraph);

            // Save the document to disk.
            const string docPath = "sample.docx";
            doc.Save(docPath);

            // Load the document back for inspection.
            Document loadedDoc = new Document(docPath);

            // Find all repeating sections.
            List<RepeatingSectionInfo> report = new List<RepeatingSectionInfo>();
            IEnumerable<StructuredDocumentTag> repeatingSections = loadedDoc
                .GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

            foreach (StructuredDocumentTag repSection in repeatingSections)
            {
                // Gather nested content controls inside this repeating section.
                List<ControlInfo> nestedControls = repSection
                    .GetChildNodes(NodeType.StructuredDocumentTag, true)
                    .OfType<StructuredDocumentTag>()
                    .Select(sdt => new ControlInfo
                    {
                        Title = sdt.Title ?? string.Empty,
                        Tag = sdt.Tag ?? string.Empty,
                        Type = sdt.SdtType.ToString()
                    })
                    .ToList();

                // Add information about this repeating section to the report.
                report.Add(new RepeatingSectionInfo
                {
                    Title = repSection.Title ?? string.Empty,
                    Tag = repSection.Tag ?? string.Empty,
                    NestedControls = nestedControls
                });
            }

            // Serialize the inspection result to JSON.
            string json = JsonConvert.SerializeObject(report, Formatting.Indented);
            const string jsonPath = "nestedControls.json";
            File.WriteAllText(jsonPath, json);

            // Output the result to the console.
            Console.WriteLine("Nested content controls within repeating sections:");
            foreach (RepeatingSectionInfo sectionInfo in report)
            {
                Console.WriteLine($"Repeating Section - Title: {sectionInfo.Title}, Tag: {sectionInfo.Tag}");
                foreach (ControlInfo ctrl in sectionInfo.NestedControls)
                {
                    Console.WriteLine($"  Nested Control - Title: {ctrl.Title}, Tag: {ctrl.Tag}, Type: {ctrl.Type}");
                }
            }

            Console.WriteLine($"Inspection data saved to '{jsonPath}'.");
        }

        // Helper class to hold information about a nested control.
        private class ControlInfo
        {
            public string Title { get; set; } = string.Empty;
            public string Tag { get; set; } = string.Empty;
            public string Type { get; set; } = string.Empty;
        }

        // Helper class to hold information about a repeating section and its nested controls.
        private class RepeatingSectionInfo
        {
            public string Title { get; set; } = string.Empty;
            public string Tag { get; set; } = string.Empty;
            public List<ControlInfo> NestedControls { get; set; } = new List<ControlInfo>();
        }
    }
}
