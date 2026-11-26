using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

namespace ContentControlTagUpdater
{
    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a sample document with several content controls.
            Document sampleDoc = new Document();

            // Ensure the document has at least one paragraph to host inline SDTs.
            Paragraph firstParagraph = sampleDoc.FirstSection.Body.FirstParagraph;

            // Plain‑text inline content control.
            StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(sampleDoc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "custName"
            };
            plainTextSdt.RemoveAllChildren();
            plainTextSdt.AppendChild(new Run(sampleDoc, "John Doe"));
            firstParagraph.AppendChild(plainTextSdt);

            // Rich‑text block‑level content control.
            StructuredDocumentTag richTextSdt = new StructuredDocumentTag(sampleDoc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "Address",
                Tag = "addr"
            };
            Paragraph addressParagraph = new Paragraph(sampleDoc);
            addressParagraph.AppendChild(new Run(sampleDoc, "123 Main St, Anytown"));
            richTextSdt.AppendChild(addressParagraph);
            sampleDoc.FirstSection.Body.AppendChild(richTextSdt);

            // Drop‑down list inline content control.
            StructuredDocumentTag dropDownSdt = new StructuredDocumentTag(sampleDoc, SdtType.DropDownList, MarkupLevel.Inline)
            {
                Title = "Country",
                Tag = "cnt"
            };
            dropDownSdt.ListItems.Add(new SdtListItem("United States", "US"));
            dropDownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
            dropDownSdt.ListItems.Add(new SdtListItem("Mexico", "MX"));
            firstParagraph.AppendChild(dropDownSdt);

            // Save the seed document.
            const string inputPath = "input.docx";
            sampleDoc.Save(inputPath);

            // Step 2: Load the document and update tags to a standardized convention.
            Document doc = new Document(inputPath);

            // Collect all StructuredDocumentTag nodes.
            List<StructuredDocumentTag> sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>()
                .ToList();

            // Prepare a report structure.
            var report = new List<object>();

            for (int i = 0; i < sdtNodes.Count; i++)
            {
                StructuredDocumentTag sdt = sdtNodes[i];
                string originalTag = sdt.Tag ?? string.Empty;
                string title = sdt.Title ?? string.Empty;

                // Standardized tag: "tag-" + lower‑cased title with spaces replaced by hyphens.
                // If title is missing, fall back to "tag-{index}".
                string newTag = !string.IsNullOrWhiteSpace(title)
                    ? $"tag-{title.Replace(' ', '-').ToLowerInvariant()}"
                    : $"tag-{i + 1}";

                sdt.Tag = newTag;

                // Record the change for the JSON report.
                report.Add(new
                {
                    Index = i + 1,
                    Title = title,
                    OriginalTag = originalTag,
                    UpdatedTag = newTag,
                    SdtType = sdt.SdtType.ToString()
                });
            }

            // Step 3: Save the updated document.
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            // Step 4: Write a JSON report of the tag updates.
            string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
            File.WriteAllText("tags-report.json", jsonReport);
        }
    }
}
