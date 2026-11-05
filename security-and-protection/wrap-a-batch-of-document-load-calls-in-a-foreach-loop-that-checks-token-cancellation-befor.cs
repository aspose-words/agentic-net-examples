using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample documents.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsSample_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        // Prepare a list of file paths for the sample documents.
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(tempFolder, $"Sample{i}.docx");
            // Create a simple document with one paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(filePath);
            docPaths.Add(filePath);
        }

        // Set up a cancellation token source.
        CancellationTokenSource cts = new CancellationTokenSource();

        // Example: cancel after the second document is loaded.
        int loadCount = 0;

        // Load each document, checking for cancellation before each load.
        foreach (string path in docPaths)
        {
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine("Loading cancelled before processing: " + path);
                break;
            }

            // Load the document.
            Document loadedDoc = new Document(path);
            loadCount++;

            Console.WriteLine($"Loaded document: {Path.GetFileName(path)} (Paragraph count: {loadedDoc.GetChildNodes(NodeType.Paragraph, true).Count})");

            // Cancel after loading two documents to demonstrate the check.
            if (loadCount == 2)
            {
                cts.Cancel();
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
