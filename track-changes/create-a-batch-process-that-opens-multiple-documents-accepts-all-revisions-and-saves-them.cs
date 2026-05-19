using System;
using System.IO;
using Aspose.Words;

public class BatchAcceptRevisions
{
    public static void Main()
    {
        // Folder to store sample documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create sample documents with revisions if they do not already exist.
        for (int i = 1; i <= 3; i++)
        {
            string filePath = Path.Combine(docsFolder, $"Sample{i}.docx");
            if (!File.Exists(filePath))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Initial content (no revision).
                builder.Writeln("Original content.");

                // Start tracking revisions.
                doc.StartTrackRevisions($"Author{i}", DateTime.Now);

                // Make some changes that will be recorded as revisions.
                builder.Writeln($"Added line {i}.");
                builder.Writeln($"Another change {i}.");

                // Stop tracking.
                doc.StopTrackRevisions();

                // Save the document.
                doc.Save(filePath);
            }
        }

        // Process each document: accept all revisions and save in place.
        string[] files = Directory.GetFiles(docsFolder, "*.docx");
        foreach (string file in files)
        {
            Document doc = new Document(file);

            // Accept all tracked changes if any exist.
            if (doc.HasRevisions)
            {
                doc.AcceptAllRevisions();
            }

            // Overwrite the original file with the updated document.
            doc.Save(file);
        }

        // Optional: indicate completion.
        Console.WriteLine("All documents processed and revisions accepted.");
    }
}
