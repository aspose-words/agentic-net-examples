using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare a temporary folder for sample documents.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeWordsDemo");
        Directory.CreateDirectory(artifactsDir);

        // Create a batch of sample documents.
        var filePaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            string filePath = Path.Combine(artifactsDir, $"Doc{i}.docx");
            doc.Save(filePath);
            filePaths.Add(filePath);
        }

        // Set up a cancellation token (not cancelled in this example).
        using CancellationTokenSource cts = new CancellationTokenSource();

        // Iterate over the batch, checking cancellation before each load.
        foreach (string path in filePaths)
        {
            if (cts.Token.IsCancellationRequested)
            {
                Console.WriteLine("Loading operation was cancelled.");
                break;
            }

            // Load the document.
            Document loadedDoc = new Document(path);
            Console.WriteLine($"Loaded '{Path.GetFileName(path)}' with text: {loadedDoc.GetText().Trim()}");
        }

        // Clean up temporary files (optional).
        foreach (string path in filePaths)
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        Directory.Delete(artifactsDir, true);
    }
}
