using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample documents.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempFolder);

        // Create sample documents with numbered lists.
        CreateSampleDocument(Path.Combine(tempFolder, "Doc1.docx"));
        CreateSampleDocument(Path.Combine(tempFolder, "Doc2.docx"));

        // Load all documents from the folder into a collection.
        List<Document> documents = new List<Document>();
        foreach (string filePath in Directory.GetFiles(tempFolder, "*.docx"))
        {
            documents.Add(new Document(filePath));
        }

        // Iterate through each document and convert numbered lists to bulleted lists.
        foreach (Document doc in documents)
        {
            // Find all paragraphs that are part of a list.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.IsListItem)
                {
                    // Convert the current list item to a bulleted list.
                    para.ListFormat.ApplyBulletDefault();
                }
            }

            // Save the modified document.
            string outputFileName = "converted_" + Path.GetFileName(doc.OriginalFileName ?? "output.docx");
            string outputPath = Path.Combine(tempFolder, outputFileName);
            doc.Save(outputPath);
        }

        // The program ends here; no console output is required.
    }

    // Helper method to create a sample document containing a numbered list.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a numbered list with two items.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First numbered item");
        builder.Writeln("Second numbered item");
        builder.ListFormat.RemoveNumbers(); // End of list.

        doc.Save(filePath);
    }
}
