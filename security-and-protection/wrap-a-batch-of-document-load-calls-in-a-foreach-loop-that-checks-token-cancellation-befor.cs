using System;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample documents and generate them.
        string docsDir = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsDir);
        CreateSampleDocuments(docsDir);

        // Gather all .docx files to load.
        string[] docFiles = Directory.GetFiles(docsDir, "*.docx");

        // Set up a cancellation token that will be triggered after a short delay.
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(200); // milliseconds

        // Folder for processed copies.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Load each document, checking for cancellation before each load.
        foreach (string filePath in docFiles)
        {
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine($"Loading cancelled before processing: {Path.GetFileName(filePath)}");
                break;
            }

            // Load the document.
            Document doc = new Document(filePath);

            // Save a copy to the output folder.
            string outPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outPath);

            Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
        }

        Console.WriteLine("Finished.");
    }

    // Helper method to create a few simple documents.
    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 5; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i}");
            string path = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(path);
        }
    }
}
