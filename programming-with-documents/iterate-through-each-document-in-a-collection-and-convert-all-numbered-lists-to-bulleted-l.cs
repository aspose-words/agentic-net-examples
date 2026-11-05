using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for sample documents.
        string folderPath = Path.Combine(Path.GetTempPath(), "AsposeDocsDemo");
        Directory.CreateDirectory(folderPath);

        // Create sample documents containing numbered lists.
        CreateSampleDocument(Path.Combine(folderPath, "Doc1.docx"));
        CreateSampleDocument(Path.Combine(folderPath, "Doc2.docx"));

        // Get all .docx files in the folder.
        string[] docFiles = Directory.GetFiles(folderPath, "*.docx");

        // Process each document: convert numbered lists to bulleted lists.
        foreach (string filePath in docFiles)
        {
            Document doc = new Document(filePath);

            // Iterate through all paragraphs.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Check if the paragraph is part of a list.
                if (para.IsListItem)
                {
                    // Preserve the original list level.
                    int originalLevel = para.ListFormat.ListLevelNumber;

                    // Apply a bulleted list style to the paragraph.
                    para.ListFormat.ApplyBulletDefault();

                    // Restore the original level (if it was deeper than level 1).
                    para.ListFormat.ListLevelNumber = originalLevel;
                }
            }

            // Save the modified document with a new name.
            string outputPath = Path.Combine(folderPath, $"Output_{Path.GetFileName(filePath)}");
            doc.Save(outputPath);
        }

        // No console output required; the output files are created in the temporary folder.
    }

    // Helper method to create a sample document with a numbered list.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample numbered list:");
        builder.ListFormat.ApplyNumberDefault(); // Start a numbered list.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.ListFormat.RemoveNumbers(); // End the list.

        doc.Save(filePath);
    }
}
