using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;

public class ParallelExtractionExample
{
    // Number of sample documents to create and process.
    private const int DocumentCount = 5;

    // Folder names for input and output files.
    private static readonly string InputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
    private static readonly string OutputFolder = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedTexts");

    public static void Main()
    {
        // Ensure clean input and output directories.
        PrepareDirectory(InputFolder);
        PrepareDirectory(OutputFolder);

        // Create sample documents.
        List<string> inputFiles = new List<string>();
        for (int i = 1; i <= DocumentCount; i++)
        {
            string filePath = Path.Combine(InputFolder, $"SampleDoc_{i}.docx");
            CreateSampleDocument(filePath, i);
            inputFiles.Add(filePath);
        }

        // Parallel extraction of document bodies.
        Parallel.ForEach(inputFiles, filePath =>
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Extract the text of the main body range.
            string extractedText = doc.FirstSection.Body.Range.Text;

            // Validate extraction.
            if (string.IsNullOrWhiteSpace(extractedText))
                throw new InvalidOperationException($"Extraction failed for file '{filePath}'.");

            // Determine output file name.
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string outputPath = Path.Combine(OutputFolder, $"{fileName}_Extracted.txt");

            // Write extracted text to a file.
            File.WriteAllText(outputPath, extractedText);
        });

        // Simple verification that all output files were created.
        int createdFiles = Directory.GetFiles(OutputFolder, "*_Extracted.txt").Length;
        if (createdFiles != DocumentCount)
            throw new InvalidOperationException("Not all extracted files were created.");

        // Indicate successful completion.
        Console.WriteLine($"Successfully extracted text from {DocumentCount} documents in parallel.");
    }

    // Creates a deterministic sample document with identifiable content.
    private static void CreateSampleDocument(string filePath, int index)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln($"--- Document {index} Header ---");
        builder.Writeln($"This is the first paragraph of document {index}.");
        builder.Writeln($"This is the second paragraph of document {index}.");
        builder.Writeln($"--- Document {index} Footer ---");

        // Save the document to the specified path.
        doc.Save(filePath);
    }

    // Ensures a directory exists and is empty.
    private static void PrepareDirectory(string path)
    {
        if (Directory.Exists(path))
        {
            foreach (string file in Directory.GetFiles(path))
                File.Delete(file);
        }
        else
        {
            Directory.CreateDirectory(path);
        }
    }
}
