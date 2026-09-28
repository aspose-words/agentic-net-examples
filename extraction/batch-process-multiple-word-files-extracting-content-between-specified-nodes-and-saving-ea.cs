using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class BatchWordToPdfExtractor
{
    public static void Main()
    {
        // Prepare working directories
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample Word documents
        CreateSampleDocument(Path.Combine(inputDir, "doc1.docx"), "Document One");
        CreateSampleDocument(Path.Combine(inputDir, "doc2.docx"), "Document Two");
        CreateSampleDocument(Path.Combine(inputDir, "doc3.docx"), "Document Three");

        // Process each document
        var report = new List<object>();
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string pdfPath = Path.Combine(outputDir, fileName + ".pdf");

            Document source = new Document(filePath);

            // Locate start and end bookmarks
            Bookmark startBookmark = source.Range.Bookmarks["Start"];
            Bookmark endBookmark = source.Range.Bookmarks["End"];
            if (startBookmark == null || endBookmark == null)
                throw new InvalidOperationException($"Bookmarks not found in {fileName}.");

            // Determine the paragraphs that contain the bookmarks
            Paragraph startPara = startBookmark.BookmarkStart.ParentNode as Paragraph;
            Paragraph endPara = endBookmark.BookmarkEnd.ParentNode as Paragraph;
            if (startPara == null || endPara == null)
                throw new InvalidOperationException($"Bookmark containers not found in {fileName}.");

            // Build a new document with the extracted range
            Document extracted = new Document();
            extracted.RemoveAllChildren();
            Section section = new Section(extracted);
            extracted.AppendChild(section);
            Body body = new Body(extracted);
            section.AppendChild(body);

            // Use NodeImporter to import nodes from source to extracted document
            NodeImporter importer = new NodeImporter(source, extracted, ImportFormatMode.KeepSourceFormatting);

            bool copying = false;
            foreach (Node node in source.FirstSection.Body.GetChildNodes(NodeType.Any, false))
            {
                if (node == startPara)
                    copying = true;

                if (copying)
                {
                    Node importedNode = importer.ImportNode(node, true);
                    body.AppendChild(importedNode);
                }

                if (node == endPara)
                    break;
            }

            // Save the extracted content as PDF
            extracted.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for {fileName}.");

            report.Add(new { Document = fileName, PdfPath = pdfPath });
        }

        // Write a JSON report
        string jsonReportPath = Path.Combine(baseDir, "extraction_report.json");
        File.WriteAllText(jsonReportPath, JsonConvert.SerializeObject(report, Formatting.Indented));

        // Ensure the report was written
        if (!File.Exists(jsonReportPath))
            throw new InvalidOperationException("JSON report was not created.");
    }

    private static void CreateSampleDocument(string path, string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.Writeln("Intro paragraph before the extraction range.");

        // Start bookmark
        builder.StartBookmark("Start");
        builder.Writeln("This paragraph is inside the extraction range.");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("Another paragraph inside the extraction range.");
        builder.EndBookmark("Start");

        // End bookmark (separate to demonstrate range)
        builder.StartBookmark("End");
        builder.Writeln("Final paragraph inside the extraction range.");
        builder.EndBookmark("End");

        builder.Writeln("Paragraph after the extraction range.");

        doc.Save(path);
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample document at {path}.");
    }
}
