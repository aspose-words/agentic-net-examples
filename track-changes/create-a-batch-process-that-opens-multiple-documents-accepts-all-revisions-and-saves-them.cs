using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class BatchRevisionProcessor
{
    public static void Main()
    {
        // Folder to store sample documents.
        string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(docsFolder);

        // Create sample documents with revisions.
        CreateSampleDocuments(docsFolder, 3);

        // Process each document: accept all revisions and save in place.
        foreach (string filePath in Directory.GetFiles(docsFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            if (doc.HasRevisions)
            {
                doc.AcceptAllRevisions();
                doc.Save(filePath);
            }
        }
    }

    private static void CreateSampleDocuments(string folderPath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            string fileName = $"Sample{i}.docx";
            string filePath = Path.Combine(folderPath, fileName);

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start tracking revisions.
            doc.StartTrackRevisions($"Author{i}", DateTime.Now);

            // Add some text that will be recorded as revisions.
            builder.Writeln($"This is the first paragraph of document {i}.");
            builder.Writeln($"Adding a second paragraph to document {i}.");

            // Stop tracking revisions.
            doc.StopTrackRevisions();

            // Save the document.
            doc.Save(filePath);
        }
    }
}
