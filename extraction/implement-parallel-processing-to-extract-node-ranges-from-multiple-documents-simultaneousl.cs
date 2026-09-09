using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Set up input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample source documents.
        const int documentCount = 5;
        List<string> inputFiles = new List<string>();
        for (int i = 1; i <= documentCount; i++)
        {
            string filePath = Path.Combine(inputDir, $"Sample_{i}.docx");
            CreateSampleDocument(filePath, i);
            inputFiles.Add(filePath);
        }

        // Process each document in parallel.
        Parallel.ForEach(inputFiles, inputFile =>
        {
            // Load the source document.
            Document srcDoc = new Document(inputFile);

            // Get the first paragraph.
            Paragraph firstParagraph = srcDoc.FirstSection?.Body?.Paragraphs?[0];
            if (firstParagraph == null)
                throw new InvalidOperationException($"Document '{inputFile}' does not contain any paragraphs.");

            // Prepare a new document that will hold the extracted paragraph.
            Document extractedDoc = new Document();
            extractedDoc.RemoveAllChildren();

            Section section = new Section(extractedDoc);
            extractedDoc.AppendChild(section);

            Body body = new Body(extractedDoc);
            section.AppendChild(body);

            // Import the paragraph from the source document into the destination document.
            NodeImporter importer = new NodeImporter(srcDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedNode = importer.ImportNode(firstParagraph, true);
            body.AppendChild(importedNode);

            // Save the extracted document.
            string fileName = Path.GetFileNameWithoutExtension(inputFile);
            string outputPath = Path.Combine(outputDir, $"{fileName}_Extracted.docx");
            extractedDoc.Save(outputPath);
        });

        // Verify that all output files were created.
        string[] resultFiles = Directory.GetFiles(outputDir, "*_Extracted.docx");
        if (resultFiles.Length != documentCount)
            throw new InvalidOperationException("Not all extracted documents were generated.");

        // Write a JSON report of the processed files.
        var report = new
        {
            ProcessedAt = DateTime.UtcNow,
            InputFiles = inputFiles,
            OutputFiles = resultFiles
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(Path.Combine(baseDir, "ExtractionReport.json"), jsonReport);
    }

    // Creates a sample document with identifiable content.
    private static void CreateSampleDocument(string filePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln($"Document #{index} - Introduction");
        builder.Writeln($"This is the first paragraph of document {index}.");
        builder.Writeln($"Additional content for document {index}.");

        doc.Save(filePath);
    }
}
