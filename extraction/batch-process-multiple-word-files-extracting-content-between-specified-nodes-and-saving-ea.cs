using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class BatchWordExtraction
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = AppContext.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample Word documents
        CreateSampleDocuments(inputFolder);

        // Process each document
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            ProcessDocument(filePath, outputFolder);
        }

        Console.WriteLine("Extraction completed. PDFs are located in: " + outputFolder);
    }

    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Intro paragraph
            builder.Writeln($"Document {i} - Introduction.");

            // Start bookmark (placed on its own paragraph)
            builder.StartBookmark("Start");
            builder.Writeln($"Start of extracted content in document {i}.");
            builder.Writeln($"Another paragraph inside the extracted range of document {i}.");
            builder.EndBookmark("Start");

            // End bookmark (placed on its own paragraph)
            builder.StartBookmark("End");
            builder.Writeln($"End of extracted content in document {i}.");
            builder.EndBookmark("End");

            // Closing paragraph
            builder.Writeln($"Document {i} - Conclusion.");

            string fileName = Path.Combine(folder, $"Doc{i}.docx");
            doc.Save(fileName);
        }
    }

    private static void ProcessDocument(string inputPath, string outputFolder)
    {
        Document srcDoc = new Document(inputPath);

        // Validate bookmarks exist
        Bookmark startBookmark = srcDoc.Range.Bookmarks["Start"];
        Bookmark endBookmark = srcDoc.Range.Bookmarks["End"];
        if (startBookmark == null || endBookmark == null)
        {
            throw new InvalidOperationException($"Required bookmarks not found in {Path.GetFileName(inputPath)}.");
        }

        // Determine the paragraphs that bound the bookmarks
        Paragraph startParagraph = startBookmark.BookmarkStart.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkEnd.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
        {
            throw new InvalidOperationException("Bookmarks are not placed inside paragraphs.");
        }

        // Collect paragraphs between the start and end bookmarks (inclusive)
        List<Paragraph> extractedParagraphs = new List<Paragraph>();
        NodeCollection allParagraphs = srcDoc.GetChildNodes(NodeType.Paragraph, true);
        bool withinRange = false;
        foreach (Paragraph para in allParagraphs)
        {
            if (para == startParagraph)
            {
                withinRange = true;
            }

            if (withinRange)
            {
                extractedParagraphs.Add(para);
            }

            if (para == endParagraph)
            {
                break;
            }
        }

        if (extractedParagraphs.Count == 0)
        {
            throw new InvalidOperationException($"No content extracted from {Path.GetFileName(inputPath)}.");
        }

        // Create a new document for the extracted content
        Document destDoc = new Document();
        destDoc.RemoveAllChildren();

        // Add a new section with a body
        Section section = new Section(destDoc);
        Body body = new Body(destDoc);
        section.AppendChild(body);
        destDoc.AppendChild(section);

        // Import and append extracted paragraphs
        foreach (Paragraph para in extractedParagraphs)
        {
            Node imported = destDoc.ImportNode(para, true, ImportFormatMode.KeepSourceFormatting);
            body.AppendChild(imported);
        }

        // Save as PDF
        string outputFileName = Path.GetFileNameWithoutExtension(inputPath) + "_Extracted.pdf";
        string outputPath = Path.Combine(outputFolder, outputFileName);
        destDoc.Save(outputPath, SaveFormat.Pdf);
    }
}
