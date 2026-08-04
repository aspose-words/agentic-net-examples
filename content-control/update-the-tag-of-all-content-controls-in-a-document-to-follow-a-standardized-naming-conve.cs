using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json; // Included as per required packages

namespace ContentControlTagUpdater
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the sample input and output documents.
            const string inputPath = "input.docx";
            const string outputPath = "output.docx";

            // -----------------------------------------------------------------
            // Step 1: Create a sample document with several content controls.
            // -----------------------------------------------------------------
            Document seedDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(seedDoc);

            // Plain‑text content control.
            StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(seedDoc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "old-tag-plain"
            };
            plainTextSdt.RemoveAllChildren();
            plainTextSdt.AppendChild(new Run(seedDoc, "Alice"));
            builder.InsertNode(plainTextSdt);
            builder.Writeln(); // Move to next line.

            // Rich‑text (block‑level) content control.
            StructuredDocumentTag richTextSdt = new StructuredDocumentTag(seedDoc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "AddressBlock",
                Tag = "old-tag-rich"
            };
            Paragraph para = new Paragraph(seedDoc);
            para.AppendChild(new Run(seedDoc, "123 Main St, Springfield"));
            richTextSdt.AppendChild(para);
            seedDoc.FirstSection.Body.AppendChild(richTextSdt);
            builder.Writeln(); // Ensure separation.

            // Checkbox content control.
            StructuredDocumentTag checkboxSdt = new StructuredDocumentTag(seedDoc, SdtType.Checkbox, MarkupLevel.Inline)
            {
                Title = "Subscribe",
                Tag = "old-tag-checkbox",
                Checked = false
            };
            builder.InsertNode(checkboxSdt);
            builder.Writeln();

            // Save the seed document.
            seedDoc.Save(inputPath);

            // -----------------------------------------------------------------
            // Step 2: Load the document and update all content control tags.
            // -----------------------------------------------------------------
            Document doc = new Document(inputPath);

            // Retrieve all StructuredDocumentTag nodes in the document.
            var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                              .OfType<StructuredDocumentTag>()
                              .ToList();

            // Apply a standardized naming convention: "Tag_1", "Tag_2", ...
            int index = 1;
            foreach (var sdt in sdtNodes)
            {
                sdt.Tag = $"Tag_{index}";
                index++;
            }

            // -----------------------------------------------------------------
            // Step 3: Save the updated document.
            // -----------------------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
