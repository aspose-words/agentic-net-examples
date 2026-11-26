using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace ExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample DOCM file containing a macro button field,
            //    some paragraphs to be extracted, and an end‑marker paragraph.
            // -----------------------------------------------------------------
            string sourcePath = "sample.docm";
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Paragraph before the macro field.");

            // Insert a macro button field (macro‑enabled field).
            builder.InsertField("MACROBUTTON NoMacro ClickMe", "ClickMe");

            // Content that will be extracted (between the field and the end marker).
            builder.Writeln("Extracted paragraph 1.");
            builder.Writeln("Extracted paragraph 2.");

            // End marker paragraph.
            builder.Writeln("End marker paragraph.");

            // Additional content after the end marker (not to be extracted).
            builder.Writeln("Paragraph after the end marker.");

            // Save the sample document as DOCM.
            sourceDoc.Save(sourcePath, SaveFormat.Docm);

            // -----------------------------------------------------------------
            // 2. Load the DOCM file and locate the macro button field and the
            //    end‑marker paragraph.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(sourcePath);

            Field macroField = loadedDoc.Range.Fields
                .FirstOrDefault(f => f.Type == FieldType.FieldMacroButton);
            if (macroField == null)
                throw new InvalidOperationException("Macro button field not found.");

            Paragraph endParagraph = loadedDoc.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .FirstOrDefault(p => p.GetText().Trim() == "End marker paragraph.");
            if (endParagraph == null)
                throw new InvalidOperationException("End marker paragraph not found.");

            // -----------------------------------------------------------------
            // 3. Build a new document that will contain the extracted content.
            // -----------------------------------------------------------------
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren();

            Section resultSection = new Section(resultDoc);
            resultDoc.AppendChild(resultSection);

            Body resultBody = new Body(resultDoc);
            resultSection.AppendChild(resultBody);

            // -----------------------------------------------------------------
            // 4. Import nodes that lie between the field end and the end paragraph.
            // -----------------------------------------------------------------
            NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

            // Start with the node immediately after the field end.
            Node current = macroField.End.NextSibling;

            while (current != null && current != endParagraph)
            {
                // Only block‑level nodes (Paragraph, Table, etc.) can be appended directly to Body.
                if (current.NodeType == NodeType.Paragraph || current.NodeType == NodeType.Table)
                {
                    Node importedNode = importer.ImportNode(current, true);
                    resultBody.AppendChild(importedNode);
                }

                // Move to the next sibling in the source document.
                current = current.NextSibling;
            }

            // -----------------------------------------------------------------
            // 5. Save the extracted content as DOCX and verify the file was created.
            // -----------------------------------------------------------------
            string outputPath = "extracted.docx";
            resultDoc.Save(outputPath, SaveFormat.Docx);

            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The extracted DOCX file was not created.");
        }
    }
}
