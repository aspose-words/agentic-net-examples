using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Directory to store temporary sample documents.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsBatchDemo");
        Directory.CreateDirectory(tempDir);

        // Create a list of sample document file paths.
        List<string> docPaths = new List<string>();
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(tempDir, $"SampleDocument{i}.docx");
            CreateSampleDocument(filePath, $"Author{i}");
            docPaths.Add(filePath);
        }

        // Batch process: open each document, accept all revisions, and save in place.
        foreach (string path in docPaths)
        {
            Document doc = new Document(path);
            // Ensure there are revisions before accepting.
            if (doc.HasRevisions)
            {
                doc.AcceptAllRevisions();
                // After acceptance, there should be no revisions left.
                if (doc.Revisions.Count != 0)
                {
                    throw new InvalidOperationException($"Revisions were not fully accepted in {path}.");
                }
            }
            // Save the document back to the same file (in place).
            doc.Save(path);
        }

        // Optional: indicate processing is complete.
        Console.WriteLine($"Processed {docPaths.Count} documents. Files are located at: {tempDir}");
    }

    // Creates a sample document with tracked changes.
    private static void CreateSampleDocument(string filePath, string author)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Initial content (no revisions).
        builder.Writeln("This is the original content of the document.");

        // Start tracking revisions.
        doc.StartTrackRevisions(author, DateTime.Now);

        // Add some revisions.
        builder.Writeln($"Inserted line by {author}.");
        builder.Writeln($"Another inserted line by {author}.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document.
        doc.Save(filePath);
    }
}
