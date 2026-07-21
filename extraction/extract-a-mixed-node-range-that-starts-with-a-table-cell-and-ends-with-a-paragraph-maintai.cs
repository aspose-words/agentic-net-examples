using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // -------------------------------------------------
            // 1. Create a sample source document.
            // -------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Intro paragraph before the table.
            builder.Writeln("Intro paragraph before the table.");

            // Build a 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1, Row 1");
            builder.InsertCell();
            builder.Write("Cell 2, Row 1");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 1, Row 2");
            builder.InsertCell();
            builder.Write("Cell 2, Row 2");
            builder.EndRow();
            builder.EndTable();

            // Paragraph that will serve as the end boundary of the extraction.
            builder.Writeln("Paragraph after the table – end of extraction range.");

            // Save the source document (so we have a physical file to load later).
            const string sourcePath = "source.docx";
            sourceDoc.Save(sourcePath);

            // -------------------------------------------------
            // 2. Load the document back (demonstrates load workflow).
            // -------------------------------------------------
            Document loadedDoc = new Document(sourcePath);

            // -------------------------------------------------
            // 3. Locate the start node (first cell of the first table) 
            //    and the end node (the paragraph after the table).
            // -------------------------------------------------
            Cell startCell = loadedDoc.GetChildNodes(NodeType.Cell, true)[0] as Cell;
            if (startCell == null)
                throw new InvalidOperationException("Start cell not found.");

            Paragraph endParagraph = null;
            foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.GetText().Contains("Paragraph after the table"))
                {
                    endParagraph = para;
                    break;
                }
            }
            if (endParagraph == null)
                throw new InvalidOperationException("End paragraph not found.");

            // -------------------------------------------------
            // 4. Prepare the destination document.
            // -------------------------------------------------
            Document extractedDoc = new Document();
            extractedDoc.RemoveAllChildren();

            Section section = new Section(extractedDoc);
            extractedDoc.AppendChild(section);
            Body body = new Body(extractedDoc);
            section.AppendChild(body);

            // -------------------------------------------------
            // 5. Import nodes from the source document to the destination.
            // -------------------------------------------------
            // NodeImporter handles the required import (style, list, etc.) and creates nodes that belong to the destination document.
            NodeImporter importer = new NodeImporter(loadedDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

            // Start traversal from the start cell.
            Node currentNode = startCell;
            while (currentNode != null)
            {
                // Skip Cell nodes – the table will be imported as a whole later.
                if (currentNode.NodeType == NodeType.Cell)
                {
                    // Do nothing; the Table node will be processed later.
                }
                else if (currentNode.NodeType == NodeType.Paragraph || currentNode.NodeType == NodeType.Table)
                {
                    // Import the node into the destination document.
                    Node importedNode = importer.ImportNode(currentNode, true);
                    // Append block‑level nodes directly to the body.
                    body.AppendChild(importedNode);
                }

                // Stop when we have processed the end paragraph.
                if (currentNode == endParagraph)
                    break;

                // Move to the next node in document order.
                currentNode = currentNode.NextPreOrder(loadedDoc);
            }

            // -------------------------------------------------
            // 6. Save the extracted content.
            // -------------------------------------------------
            const string extractedPath = "extracted.docx";
            extractedDoc.Save(extractedPath);

            // Verify that the output file was created.
            if (!File.Exists(extractedPath))
                throw new InvalidOperationException("The extracted document was not created.");

            Console.WriteLine($"Extraction completed. Output saved to '{extractedPath}'.");
        }
    }
}
