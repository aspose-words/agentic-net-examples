using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ParallelExtractionExample
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents.
        int documentCount = 5;
        List<string> inputFiles = new List<string>();
        for (int i = 1; i <= documentCount; i++)
        {
            string fileName = $"SampleDocument{i}.docx";
            string filePath = Path.Combine(inputDir, fileName);
            CreateSampleDocument(filePath, i);
            inputFiles.Add(filePath);
        }

        // Extract first two paragraphs from each document in parallel.
        List<string> extractedFiles = new List<string>();
        Parallel.ForEach(inputFiles, inputPath =>
        {
            string outputFileName = Path.GetFileNameWithoutExtension(inputPath) + "_Extracted.docx";
            string outputPath = Path.Combine(outputDir, outputFileName);

            ExtractFirstTwoParagraphs(inputPath, outputPath);

            lock (extractedFiles)
            {
                extractedFiles.Add(outputPath);
            }
        });

        // Write JSON report.
        string reportPath = Path.Combine(outputDir, "ExtractionReport.json");
        var report = new
        {
            Timestamp = DateTime.UtcNow,
            ExtractedFiles = extractedFiles
        };
        File.WriteAllText(reportPath, JsonConvert.SerializeObject(report, Formatting.Indented));

        // Validate report creation.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Extraction report was not generated.");
    }

    // Creates a simple DOCX file with a few paragraphs.
    private static void CreateSampleDocument(string filePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln($"Document {index} - Introduction.");
        builder.Writeln($"Document {index} - Body paragraph one.");
        builder.Writeln($"Document {index} - Body paragraph two.");
        builder.Writeln($"Document {index} - Conclusion.");

        doc.Save(filePath);

        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample document: {filePath}");
    }

    // Extracts the first two paragraphs from the source document and saves them as a new document.
    private static void ExtractFirstTwoParagraphs(string sourcePath, string destinationPath)
    {
        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Ensure there are at least two paragraphs.
        ParagraphCollection paragraphs = sourceDoc.FirstSection.Body.Paragraphs;
        if (paragraphs.Count < 2)
            throw new InvalidOperationException("Source document does not contain enough paragraphs for extraction.");

        Paragraph firstParagraph = paragraphs[0];
        Paragraph secondParagraph = paragraphs[1];

        // Create a new empty document with a proper structure.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section section = new Section(resultDoc);
        resultDoc.AppendChild(section);

        Body body = new Body(resultDoc);
        section.AppendChild(body);

        // Import the selected paragraphs into the new document.
        NodeImporter importer = new NodeImporter(sourceDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedFirst = importer.ImportNode(firstParagraph, true);
        Node importedSecond = importer.ImportNode(secondParagraph, true);

        body.AppendChild(importedFirst);
        body.AppendChild(importedSecond);

        // Save the extracted content.
        resultDoc.Save(destinationPath);

        if (!File.Exists(destinationPath))
            throw new InvalidOperationException($"Failed to save extracted document: {destinationPath}");
    }
}
