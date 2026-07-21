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
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempDir);

        // Create a few sample documents.
        var sampleFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(tempDir, $"Sample{i}.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            doc.Save(filePath);
            sampleFiles.Add(filePath);
        }

        // Set up a cancellation token source.
        var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Load each document, checking for cancellation before each load.
        foreach (string file in sampleFiles)
        {
            if (token.IsCancellationRequested)
            {
                Console.WriteLine("Cancellation requested. Stopping further loads.");
                break;
            }

            // Load the document.
            var loadedDoc = new Document(file);
            Console.WriteLine($"Loaded '{Path.GetFileName(file)}' with text: {loadedDoc.GetText().Trim()}");

            // For demonstration, cancel after the first document is loaded.
            if (file.EndsWith("Sample1.docx"))
            {
                cts.Cancel();
            }
        }

        // Clean up temporary files.
        foreach (string file in sampleFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
        Directory.Delete(tempDir, true);
    }
}
