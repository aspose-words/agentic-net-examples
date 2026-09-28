using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

namespace ExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample source document containing a table, an image,
            //    and a DATE field.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Intro paragraph.
            builder.Writeln("Intro paragraph before the range.");

            // Insert a table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Insert an image (1x1 transparent PNG).
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9yhl4AAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            using (MemoryStream imageStream = new MemoryStream(pngBytes))
            {
                builder.InsertImage(imageStream);
            }

            // Insert a DATE field.
            builder.InsertField(FieldType.FieldDate, true);

            // Paragraph after the range.
            builder.Writeln("Paragraph after the range.");

            // Save the source document.
            const string sourcePath = "source.docx";
            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Load the document and locate the start and end nodes.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(sourcePath);
            Body sourceBody = loadedDoc.FirstSection.Body;

            // Start node: first table in the document.
            Table startTable = loadedDoc.GetChildNodes(NodeType.Table, true)
                                        .OfType<Table>()
                                        .FirstOrDefault();

            if (startTable == null)
                throw new InvalidOperationException("No table found to serve as start of extraction range.");

            // End node: paragraph that contains the first field.
            Field firstField = loadedDoc.Range.Fields.FirstOrDefault();

            if (firstField == null)
                throw new InvalidOperationException("No field found to serve as end of extraction range.");

            Paragraph endParagraph = firstField.Start.ParentNode as Paragraph;

            if (endParagraph == null)
                throw new InvalidOperationException("Field is not inside a paragraph.");

            // Ensure both nodes belong to the same body.
            if (startTable.ParentNode != sourceBody || endParagraph.ParentNode != sourceBody)
                throw new InvalidOperationException("Start or end node is not a direct child of the document body.");

            // -----------------------------------------------------------------
            // 3. Build a new document that will contain the extracted range.
            // -----------------------------------------------------------------
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren();

            Section resultSection = new Section(resultDoc);
            resultDoc.AppendChild(resultSection);

            Body resultBody = new Body(resultDoc);
            resultSection.AppendChild(resultBody);

            // -----------------------------------------------------------------
            // 4. Clone nodes from start to end (inclusive) preserving order.
            // -----------------------------------------------------------------
            bool withinRange = false;
            NodeCollection sourceChildren = sourceBody.GetChildNodes(NodeType.Any, true);

            foreach (Node child in sourceChildren)
            {
                if (child == startTable)
                    withinRange = true;

                if (withinRange)
                {
                    // Clone the node deeply to keep its internal structure.
                    resultBody.AppendChild(child.Clone(true));
                }

                if (child == endParagraph)
                {
                    // End of range reached.
                    break;
                }
            }

            // -----------------------------------------------------------------
            // 5. Save the extracted content.
            // -----------------------------------------------------------------
            const string outputPath = "extracted.docx";
            resultDoc.Save(outputPath);

            // Validate that the output file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Extraction failed: output file was not created.");

            // -----------------------------------------------------------------
            // 6. Write a simple JSON report confirming success.
            // -----------------------------------------------------------------
            var report = new
            {
                SourceDocument = Path.GetFullPath(sourcePath),
                ExtractedDocument = Path.GetFullPath(outputPath),
                ExtractedNodeCount = resultBody.GetChildNodes(NodeType.Any, true).Count
            };

            File.WriteAllText("extraction-report.json",
                JsonConvert.SerializeObject(report, Formatting.Indented));
        }
    }
}
