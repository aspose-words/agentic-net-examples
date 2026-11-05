using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class BatchWordExtraction
{
    public static void Main()
    {
        // Set up directories
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample Word documents with a bookmark named "Extract"
        for (int i = 0; i < 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Document {i} - Intro paragraph.");

            // Start bookmark
            builder.StartBookmark("Extract");
            builder.Writeln($"Document {i} - Extracted paragraph 1.");
            builder.Writeln($"Document {i} - Extracted paragraph 2.");

            // Insert a simple table inside the bookmark
            builder.StartTable();
            builder.InsertCell();
            builder.Write($"Cell A{i}");
            builder.InsertCell();
            builder.Write($"Cell B{i}");
            builder.EndRow();
            builder.EndTable();

            // End bookmark
            builder.EndBookmark("Extract");

            builder.Writeln($"Document {i} - Closing paragraph.");

            string docPath = Path.Combine(inputDir, $"SampleDoc{i}.docx");
            doc.Save(docPath);
            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to create sample document: {docPath}");
        }

        // Process each document: extract content inside the "Extract" bookmark and save as PDF
        List<string> generatedPdfs = new List<string>();
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string docFile in docFiles)
        {
            Document sourceDoc = new Document(docFile);
            Bookmark extractBookmark = sourceDoc.Range.Bookmarks["Extract"];
            if (extractBookmark == null)
                throw new InvalidOperationException($"Bookmark 'Extract' not found in document {docFile}.");

            Node startNode = extractBookmark.BookmarkStart;
            Node endNode = extractBookmark.BookmarkEnd;

            // Prepare a new document to hold the extracted content
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren();

            Section section = new Section(resultDoc);
            resultDoc.AppendChild(section);

            Body body = new Body(resultDoc);
            section.AppendChild(body);

            // Importer to copy nodes from source to result document
            NodeImporter importer = new NodeImporter(sourceDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

            // Walk through all nodes in document order and copy block nodes that lie between the bookmark boundaries
            bool inside = false;
            foreach (Node node in sourceDoc.GetChildNodes(NodeType.Any, true))
            {
                if (node == startNode)
                {
                    inside = true;
                    continue; // Skip the bookmark start node itself
                }

                if (node == endNode)
                {
                    break; // Stop when we reach the bookmark end node
                }

                if (inside && (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table))
                {
                    Node importedNode = importer.ImportNode(node, true);
                    body.AppendChild(importedNode);
                }
            }

            // Save the extracted content as PDF
            string pdfFileName = Path.GetFileNameWithoutExtension(docFile) + "_Extracted.pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            resultDoc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");

            generatedPdfs.Add(pdfPath);
        }

        // Write a simple JSON report of generated PDFs
        string reportPath = Path.Combine(baseDir, "ExtractionReport.json");
        string jsonReport = JsonConvert.SerializeObject(new { ExtractedPdfs = generatedPdfs }, Formatting.Indented);
        File.WriteAllText(reportPath, jsonReport);
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Failed to write extraction report.");

        // Optional cleanup: delete the sample input files
        // foreach (string file in docFiles) File.Delete(file);
    }
}
