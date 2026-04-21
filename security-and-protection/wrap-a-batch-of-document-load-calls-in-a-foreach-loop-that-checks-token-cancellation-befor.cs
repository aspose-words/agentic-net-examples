using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for sample documents.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(baseDir);

        // Create a few sample documents locally.
        var sampleFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(baseDir, $"doc{i}.docx");
            CreateSampleDocument(filePath, $"This is the content of document {i}.");
            sampleFiles.Add(filePath);
        }

        // Set up a cancellation token source.
        var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Example: cancel after the first document is loaded.
        bool firstLoaded = false;

        // Iterate over the document paths, checking cancellation before each load.
        foreach (string path in sampleFiles)
        {
            if (token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Stopping further loads.");
                break;
            }

            // Load the document.
            Document doc = new Document(path);
            Console.WriteLine($"Loaded document: {Path.GetFileName(path)}");

            // Optionally save a copy to demonstrate successful load.
            string outputPath = Path.Combine(baseDir, $"Loaded_{Path.GetFileName(path)}");
            doc.Save(outputPath);
            Console.WriteLine($"Saved copy as: {Path.GetFileName(outputPath)}");

            // Trigger cancellation after the first iteration for demonstration.
            if (!firstLoaded)
            {
                firstLoaded = true;
                cts.Cancel();
            }
        }

        Console.WriteLine("Processing completed.");
    }

    // Helper method to create a simple document with given text.
    private static void CreateSampleDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
