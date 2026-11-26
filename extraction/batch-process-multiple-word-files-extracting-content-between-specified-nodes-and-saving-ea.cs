using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace BatchExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare folders
            string baseDir = Directory.GetCurrentDirectory();
            string inputFolder = Path.Combine(baseDir, "InputDocs");
            string outputFolder = Path.Combine(baseDir, "OutputPdfs");
            Directory.CreateDirectory(inputFolder);
            Directory.CreateDirectory(outputFolder);

            // Create sample documents
            CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), "Document 1");
            CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), "Document 2");
            CreateSampleDocument(Path.Combine(inputFolder, "Doc3.docx"), "Document 3");

            // Process each document in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                // Load source document
                Document source = new Document(filePath);

                // Locate start and end bookmarks
                Bookmark startBookmark = source.Range.Bookmarks["Start"];
                Bookmark endBookmark = source.Range.Bookmarks["End"];
                if (startBookmark == null || endBookmark == null)
                    throw new InvalidOperationException($"Bookmarks not found in {Path.GetFileName(filePath)}.");

                // Create a new empty document for the extracted content
                Document extracted = new Document();
                extracted.RemoveAllChildren(); // Remove the default empty section

                // Build the required document structure: Section -> Body
                Section section = new Section(extracted);
                extracted.AppendChild(section);
                Body body = new Body(extracted);
                section.AppendChild(body);

                // Determine the first node after the start bookmark and the node before the end bookmark
                Node currentNode = startBookmark.BookmarkStart.NextSibling;
                Node endNode = endBookmark.BookmarkEnd.PreviousSibling;

                // Importer to copy nodes from source to destination
                NodeImporter importer = new NodeImporter(source, extracted, ImportFormatMode.KeepSourceFormatting);

                // Traverse and import nodes until we pass the end node
                while (currentNode != null)
                {
                    // Stop after processing the node that precedes the end bookmark
                    if (currentNode == endNode?.NextSibling)
                        break;

                    // Import only block-level nodes directly; wrap inline nodes (e.g., Run) into a Paragraph
                    if (currentNode.NodeType == NodeType.Paragraph || currentNode.NodeType == NodeType.Table)
                    {
                        Node importedNode = importer.ImportNode(currentNode, true);
                        body.AppendChild(importedNode);
                    }
                    else if (currentNode.NodeType == NodeType.Run)
                    {
                        // Wrap the run inside a new paragraph to satisfy the body’s block-level requirement
                        Paragraph wrapper = new Paragraph(extracted);
                        Node importedRun = importer.ImportNode(currentNode, true);
                        wrapper.AppendChild(importedRun);
                        body.AppendChild(wrapper);
                    }
                    else
                    {
                        // For any other node types, attempt to import if they are allowed under Body
                        Node importedNode = importer.ImportNode(currentNode, true);
                        if (importedNode.NodeType == NodeType.Paragraph || importedNode.NodeType == NodeType.Table)
                            body.AppendChild(importedNode);
                    }

                    if (currentNode == endNode)
                        break;

                    currentNode = currentNode.NextSibling;
                }

                // Save extracted content as PDF
                string outputPdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
                extracted.Save(outputPdfPath, SaveFormat.Pdf);

                // Validate output
                if (!File.Exists(outputPdfPath))
                    throw new InvalidOperationException($"Failed to create PDF for {Path.GetFileName(filePath)}.");
            }
        }

        private static void CreateSampleDocument(string filePath, string title)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Intro paragraph
            builder.Writeln($"{title} - Intro");

            // Start bookmark
            builder.StartBookmark("Start");
            builder.Writeln("This is the first paragraph inside the extraction range.");
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("Another paragraph inside the extraction range.");
            builder.EndBookmark("Start");

            // Content outside the range
            builder.Writeln("Content outside the extraction range.");

            // End bookmark (empty range marker after outside content)
            builder.StartBookmark("End");
            builder.EndBookmark("End");

            // Save the sample document
            doc.Save(filePath);
            if (!File.Exists(filePath))
                throw new InvalidOperationException($"Failed to create sample document {Path.GetFileName(filePath)}.");
        }
    }
}
