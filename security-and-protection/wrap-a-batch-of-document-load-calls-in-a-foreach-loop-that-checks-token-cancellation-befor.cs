using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for sample documents.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsSample");
        Directory.CreateDirectory(tempFolder);

        // Create a few simple documents to load later.
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(tempFolder, $"Sample{i}.docx");
            Document doc = new Document(); // empty document
            doc.Save(filePath);
            docPaths.Add(filePath);
        }

        // Set up a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // For demonstration, cancel after the first document is loaded.
        bool firstLoaded = false;

        // Iterate over the document paths, checking cancellation before each load.
        foreach (string path in docPaths)
        {
            // Check for cancellation request.
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Exiting load loop.");
                break;
            }

            // Load the document.
            Document loadedDoc = new Document(path);
            Console.WriteLine($"Loaded document: {Path.GetFileName(path)} (Pages: {loadedDoc.PageCount})");

            // Trigger cancellation after the first load to demonstrate the check.
            if (!firstLoaded)
            {
                firstLoaded = true;
                cts.Cancel(); // request cancellation for subsequent iterations
            }
        }

        // Clean up temporary files.
        foreach (string path in docPaths)
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        Directory.Delete(tempFolder, true);
    }
}
