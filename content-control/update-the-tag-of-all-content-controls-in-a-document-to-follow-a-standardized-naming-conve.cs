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
            // Paths for the sample input, output, and JSON mapping file.
            const string inputPath = "input.docx";
            const string outputPath = "output.docx";
            const string jsonPath = "tag-mapping.json";

            // Create a sample document containing several content controls.
            CreateSampleDocument(inputPath);

            // Load the document that contains the content controls.
            Document doc = new Document(inputPath);

            // Prepare a list to hold old and new tag information for reporting.
            var tagMappings = new List<object>();

            // Retrieve all StructuredDocumentTag nodes in the document.
            var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true)
                .OfType<StructuredDocumentTag>();

            foreach (StructuredDocumentTag sdt in sdtNodes)
            {
                // Preserve the original tag value.
                string oldTag = sdt.Tag;

                // Ensure the Title is not null; use a fallback if necessary.
                string title = string.IsNullOrEmpty(sdt.Title) ? "Untitled" : sdt.Title;

                // Build a standardized tag: prefix "Tag_" and replace spaces with underscores.
                string newTag = "Tag_" + title.Replace(' ', '_');

                // Apply the new tag to the content control.
                sdt.Tag = newTag;

                // Record the mapping for later inspection.
                tagMappings.Add(new { Title = title, OldTag = oldTag, NewTag = newTag });
            }

            // Save the modified document.
            doc.Save(outputPath);

            // Serialize the tag mapping information to a JSON file.
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(tagMappings, Formatting.Indented));
        }

        // Generates a simple DOCX file with a few different content controls.
        private static void CreateSampleDocument(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Plain text content control (inline).
            StructuredDocumentTag plain = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline);
            plain.Title = "Customer Name";
            plain.Tag = "custName";
            plain.RemoveAllChildren();
            plain.AppendChild(new Run(doc, "Alice"));
            // Insert the inline SDT into the current paragraph.
            builder.InsertNode(plain);
            builder.Writeln();

            // Rich text content control (block level).
            StructuredDocumentTag rich = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
            rich.Title = "Address";
            rich.Tag = "addr";
            rich.RemoveAllChildren();
            Paragraph richParagraph = new Paragraph(doc);
            richParagraph.AppendChild(new Run(doc, "123 Main St"));
            rich.AppendChild(richParagraph);
            // Append the block‑level SDT to the document body.
            doc.FirstSection.Body.AppendChild(rich);
            builder.Writeln();

            // Checkbox content control (inline).
            StructuredDocumentTag check = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            check.Title = "Subscribe";
            check.Tag = "subscribe";
            check.Checked = false;
            builder.InsertNode(check);
            builder.Writeln();

            // Save the sample document to the specified path.
            doc.Save(path);
        }
    }
}
