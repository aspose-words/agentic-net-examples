using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fonts;

namespace FontBatchProcessing
{
    public class Program
    {
        public static void Main()
        {
            // Folder to store sample documents.
            string docsFolder = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
            Directory.CreateDirectory(docsFolder);

            // Create a few sample documents.
            List<string> docPaths = new List<string>();
            for (int i = 1; i <= 3; i++)
            {
                string path = Path.Combine(docsFolder, $"Sample{i}.docx");
                CreateSampleDocument(path, $"This is sample document {i}.");
                docPaths.Add(path);
            }

            // Process each document: set Helvetica font for every Run.
            foreach (string path in docPaths)
            {
                Document doc = new Document(path);

                // Iterate over all Run nodes in the document.
                foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
                {
                    // Set the font name to Helvetica.
                    run.Font.Name = "Helvetica";
                }

                // Save the modified document (overwrite original).
                doc.Save(path);
            }

            // Verify that all documents exist after processing.
            foreach (string path in docPaths)
            {
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException($"Processed file not found: {path}");
                }
            }

            // All done.
        }

        // Helper method to create a simple document with one paragraph of text.
        private static void CreateSampleDocument(string filePath, string text)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln(text);
            doc.Save(filePath);
        }
    }
}
