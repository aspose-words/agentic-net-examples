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
            // Create a new blank document.
            Document doc = new Document();

            // Create a repeating section content control (block level).
            StructuredDocumentTag repeatingSection = new StructuredDocumentTag(doc, SdtType.RepeatingSection, MarkupLevel.Block)
            {
                Title = "RepeatingSection",
                Tag = "repeating-section"
            };

            // Add a block-level rich text content control inside the repeating section.
            StructuredDocumentTag innerBlockSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "InnerBlock",
                Tag = "inner-block"
            };
            Paragraph blockParagraph = new Paragraph(doc);
            blockParagraph.AppendChild(new Run(doc, "Block level nested content control."));
            innerBlockSdt.AppendChild(blockParagraph);
            repeatingSection.AppendChild(innerBlockSdt);

            // Add an inline plain text content control inside the repeating section.
            Paragraph inlineParagraph = new Paragraph(doc);
            StructuredDocumentTag innerInlineSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "InnerInline",
                Tag = "inner-inline"
            };
            innerInlineSdt.RemoveAllChildren();
            innerInlineSdt.AppendChild(new Run(doc, "Inline nested content control."));
            inlineParagraph.AppendChild(innerInlineSdt);
            repeatingSection.AppendChild(inlineParagraph);

            // Insert the repeating section into the document body.
            doc.FirstSection.Body.AppendChild(repeatingSection);

            // Save the sample document.
            doc.Save("sample.docx");

            // Detect nested content controls within each repeating section.
            List<object> nestedInfo = new List<object>();

            IEnumerable<StructuredDocumentTag> repeatingSections = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .Where(sdt => sdt.SdtType == SdtType.RepeatingSection);

            foreach (StructuredDocumentTag repeating in repeatingSections)
            {
                // Find all descendant content controls inside the repeating section.
                IEnumerable<StructuredDocumentTag> nestedControls = repeating.GetChildNodes(NodeType.StructuredDocumentTag, true)
                    .OfType<StructuredDocumentTag>();

                foreach (StructuredDocumentTag nested in nestedControls)
                {
                    nestedInfo.Add(new
                    {
                        ParentRepeatingTitle = repeating.Title,
                        NestedTitle = nested.Title,
                        NestedTag = nested.Tag,
                        NestedType = nested.SdtType.ToString(),
                        NestedText = nested.GetText().Trim()
                    });
                }
            }

            // Output results to console.
            Console.WriteLine("Nested content controls found within repeating sections:");
            foreach (var info in nestedInfo)
            {
                Console.WriteLine(JsonConvert.SerializeObject(info, Formatting.Indented));
            }

            // Serialize the information to a JSON file.
            string jsonOutput = JsonConvert.SerializeObject(nestedInfo, Formatting.Indented);
            File.WriteAllText("nested-sdts.json", jsonOutput);
        }
    }
}
