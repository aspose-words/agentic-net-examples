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

        // Create a few sample DOCX files.
        var sampleFiles = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(tempFolder, $"Sample{i}.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(filePath);
            sampleFiles.Add(filePath);
        }

        // Set up a cancellation token (not cancelled in this example).
        using var cts = new CancellationTokenSource();
        CancellationToken token = cts.Token;

        // Batch load documents, checking cancellation before each load.
        foreach (string file in sampleFiles)
        {
            if (token.IsCancellationRequested)
            {
                Console.WriteLine("Loading operation was cancelled.");
                break;
            }

            // Load the document.
            Document loadedDoc = new Document(file);

            // Simple operation to prove the document was loaded (e.g., count sections).
            Console.WriteLine($"Loaded '{Path.GetFileName(file)}' with {loadedDoc.Sections.Count} section(s).");
        }

        // Clean up temporary files.
        foreach (string file in sampleFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
        Directory.Delete(tempFolder, true);
    }
}
