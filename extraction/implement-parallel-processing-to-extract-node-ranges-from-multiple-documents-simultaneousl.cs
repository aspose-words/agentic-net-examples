using System;
using System.IO;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;
using Newtonsoft.Json;

public class Program
{
    // Simple DTO to hold extraction results for JSON serialization.
    private class ExtractionResult
    {
        public string SourceFile { get; set; } = string.Empty;
        public string ExtractedText { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Prepare folders for input documents and extracted outputs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputTexts");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Number of sample documents to generate.
        const int documentCount = 5;

        // Create sample DOCX files.
        for (int i = 0; i < documentCount; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Document {i} - Paragraph 1");
            builder.Writeln($"Document {i} - Paragraph 2");
            builder.Writeln($"Document {i} - Paragraph 3");
            string filePath = Path.Combine(inputFolder, $"doc{i}.docx");
            doc.Save(filePath);
        }

        // Thread‑safe collection to gather results from parallel execution.
        ConcurrentBag<ExtractionResult> results = new ConcurrentBag<ExtractionResult>();

        // Process each document in parallel.
        string[] files = Directory.GetFiles(inputFolder, "*.docx");
        Parallel.ForEach(files, file =>
        {
            // Load the document.
            Document loaded = new Document(file);

            // Extract the text of the first paragraph as an example of a node range extraction.
            Paragraph firstParagraph = loaded.FirstSection?.Body?.Paragraphs?[0];
            if (firstParagraph == null)
                throw new InvalidOperationException($"No paragraph found in {file}.");

            string extractedText = firstParagraph.GetText().Trim();

            // Write the extracted text to a deterministic output file.
            string outputFileName = Path.GetFileNameWithoutExtension(file) + "_extracted.txt";
            string outputPath = Path.Combine(outputFolder, outputFileName);
            File.WriteAllText(outputPath, extractedText);

            // Store the result for later JSON reporting.
            results.Add(new ExtractionResult
            {
                SourceFile = Path.GetFileName(file),
                ExtractedText = extractedText
            });
        });

        // Validate that all expected output files were created.
        foreach (string file in files)
        {
            string expectedOutput = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(file) + "_extracted.txt");
            if (!File.Exists(expectedOutput))
                throw new InvalidOperationException($"Expected output file was not created: {expectedOutput}");
        }

        // Serialize the summary of extractions to JSON.
        List<ExtractionResult> resultList = results.ToList();
        string json = JsonConvert.SerializeObject(resultList, Formatting.Indented);
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "extraction_summary.json");
        File.WriteAllText(jsonPath, json);

        // Final validation.
        if (!File.Exists(jsonPath))
            throw new InvalidOperationException("JSON summary file was not created.");

        // The program finishes without requiring any user interaction.
    }
}
