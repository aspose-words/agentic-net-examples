using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for sample documents.
        string artifactsDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(artifactsDir);

        // Create a few sample documents.
        var sampleFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(artifactsDir, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(filePath);
            sampleFiles.Add(filePath);
        }

        // Set up a cancellation token that will trigger after a short delay.
        using (CancellationTokenSource cts = new CancellationTokenSource())
        {
            // Cancel after 500 milliseconds (adjust as needed for demonstration).
            cts.CancelAfter(500);

            // Iterate over the document files, checking cancellation before each load.
            foreach (string file in sampleFiles)
            {
                if (cts.Token.IsCancellationRequested)
                {
                    Console.WriteLine($"Loading cancelled before processing \"{Path.GetFileName(file)}\".");
                    break;
                }

                // Load the document.
                Document loadedDoc = new Document(file);
                Console.WriteLine($"Loaded \"{Path.GetFileName(file)}\" with content: \"{loadedDoc.GetText().Trim()}\"");

                // (Optional) Perform any processing here.
            }
        }

        // Clean up the temporary files.
        foreach (string file in sampleFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
        if (Directory.Exists(artifactsDir))
            Directory.Delete(artifactsDir);
    }
}
