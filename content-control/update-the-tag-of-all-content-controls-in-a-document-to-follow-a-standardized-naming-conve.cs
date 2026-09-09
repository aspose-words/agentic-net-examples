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
            Document seedDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(seedDoc);

            // Plain text content control.
            StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(seedDoc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "old-tag-plain"
            };
            plainTextSdt.RemoveAllChildren();
            plainTextSdt.AppendChild(new Run(seedDoc, "Alice"));
            builder.InsertNode(plainTextSdt);
            builder.Writeln();

            // Rich text content control.
            StructuredDocumentTag richTextSdt = new StructuredDocumentTag(seedDoc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "Address",
                Tag = "old-tag-rich"
            };
            Paragraph para = new Paragraph(seedDoc);
            para.AppendChild(new Run(seedDoc, "123 Main St"));
            richTextSdt.AppendChild(para);
            seedDoc.FirstSection.Body.AppendChild(richTextSdt);
            builder.Writeln();

            // Checkbox content control.
            StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(seedDoc, SdtType.Checkbox, MarkupLevel.Inline)
            {
                Title = "Subscribe",
                Tag = "old-tag-checkbox",
                Checked = true
            };
            builder.InsertNode(checkboxSdt);
            builder.Writeln();

            // Drop‑down list content control.
            StructuredDocumentTag dropdownSdt = new StructuredDocumentTag(seedDoc, SdtType.DropDownList, MarkupLevel.Inline)
            {
                Title = "Country",
                Tag = "old-tag-dropdown"
            };
            dropdownSdt.ListItems.Add(new SdtListItem("USA", "US"));
            dropdownSdt.ListItems.Add(new SdtListItem("Canada", "CA"));
            builder.InsertNode(dropdownSdt);
            builder.Writeln();

            // Save the seed document.
            const string inputPath = "input.docx";
            seedDoc.Save(inputPath);

            // Step 2: Load the document and update all content control tags.
            Document doc = new Document(inputPath);
            NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            var sdts = sdtNodes.OfType<StructuredDocumentTag>().ToList();

            for (int i = 0; i < sdts.Count; i++)
            {
                StructuredDocumentTag sdt = sdts[i];
                // Standardized naming: "Tag_{Title}_{Index}"
                string sanitizedTitle = string.IsNullOrWhiteSpace(sdt.Title) ? "Untitled" : sdt.Title.Replace(" ", "_");
                sdt.Tag = $"Tag_{sanitizedTitle}_{i + 1}";
            }

            // Save the updated document.
            const string outputPath = "output.docx";
            doc.Save(outputPath);

            // Optional: Export the updated tags to a JSON file for verification.
            var tagInfo = sdts.Select(s => new { s.Title, s.Tag, Type = s.SdtType.ToString() }).ToList();
            string json = JsonConvert.SerializeObject(tagInfo, Formatting.Indented);
            File.WriteAllText("updated_tags.json", json);
        }
    }
}
