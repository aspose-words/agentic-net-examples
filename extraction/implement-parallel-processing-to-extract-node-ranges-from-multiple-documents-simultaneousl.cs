using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ParallelExtractionExample
{
    public static void Main()
    {
        // Prepare directories for input and output.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractionDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample documents.
        int docCount = 5;
        for (int i = 0; i < docCount; i++)
        {
            string filePath = Path.Combine(inputDir, $"Doc{i + 1}.docx");
            CreateSampleDocument(filePath, i + 1);
        }

        // Get all document files to process.
        string[] files = Directory.GetFiles(inputDir, "*.docx");

        // Thread‑safe collection for extraction results.
        var extractionResults = new List<ExtractionResult>();
        var lockObj = new object();

        // Process each document in parallel.
        Parallel.ForEach(files, file =>
        {
            // Load the document.
            Document doc = new Document(file);

            // Extract the first paragraph's text.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            string extractedText = firstParagraph.GetText().Trim();

            // Save the extracted text to a .txt file.
            string txtFileName = Path.GetFileNameWithoutExtension(file) + "_Extracted.txt";
            string txtPath = Path.Combine(outputDir, txtFileName);
            File.WriteAllText(txtPath, extractedText);

            // Record the extraction details.
            var result = new ExtractionResult
            {
                SourceFile = Path.GetFileName(file),
                ExtractedFile = txtFileName,
                ExtractedText = extractedText
            };

            lock (lockObj)
            {
                extractionResults.Add(result);
            }
        });

        // Verify that each expected output file exists.
        foreach (string file in files)
        {
            string expectedTxt = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(file) + "_Extracted.txt");
            if (!File.Exists(expectedTxt))
                throw new InvalidOperationException($"Extraction output not found: {expectedTxt}");
        }

        // Create a JSON summary report.
        string jsonReportPath = Path.Combine(outputDir, "ExtractionReport.json");
        string json = JsonConvert.SerializeObject(extractionResults, Formatting.Indented);
        File.WriteAllText(jsonReportPath, json);

        if (!File.Exists(jsonReportPath))
            throw new InvalidOperationException("JSON report was not created.");

        // Indicate successful completion.
        Console.WriteLine("Parallel extraction completed successfully.");
    }

    // Helper to create a simple document with identifiable content.
    private static void CreateSampleDocument(string filePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"Document {index} - Introduction.");
        builder.Writeln($"Document {index} - Body paragraph one.");
        builder.Writeln($"Document {index} - Body paragraph two.");
        builder.Writeln($"Document {index} - Conclusion.");
        doc.Save(filePath);
    }

    // Simple DTO for the JSON report.
    private class ExtractionResult
    {
        public string SourceFile { get; set; }
        public string ExtractedFile { get; set; }
        public string ExtractedText { get; set; }
    }
}
