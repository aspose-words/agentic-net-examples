using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

class ConvertNumberedToBulleted
{
    static void Main()
    {
        // Prepare a temporary folder for source and output documents.
        string baseFolder = Path.Combine(Path.GetTempPath(), "DocsDemo");
        string sourceFolder = Path.Combine(baseFolder, "Source");
        string outputFolder = Path.Combine(baseFolder, "Converted");

        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents if they do not already exist.
        List<string> sourceFiles = new List<string>
        {
            Path.Combine(sourceFolder, "Doc1.docx"),
            Path.Combine(sourceFolder, "Doc2.docx")
        };

        CreateSampleDocumentIfMissing(sourceFiles[0], "First document");
        CreateSampleDocumentIfMissing(sourceFiles[1], "Second document");

        foreach (string filePath in sourceFiles)
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Create a single bulleted list that will replace all numbered lists.
            List bulletedList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Iterate through every paragraph in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in paragraphs)
            {
                // Skip paragraphs that are not part of a list.
                if (!para.ListFormat.IsListItem)
                    continue;

                // Preserve the original list level.
                int levelNumber = para.ListFormat.ListLevelNumber;

                // Retrieve the original list level definition.
                ListLevel originalLevel = para.ListFormat.List.ListLevels[levelNumber];

                // If the original list already uses bullets, leave it unchanged.
                if (originalLevel.NumberStyle == NumberStyle.Bullet)
                    continue;

                // Apply the bulleted list while keeping the same level.
                para.ListFormat.List = bulletedList;
                para.ListFormat.ListLevelNumber = levelNumber;
            }

            // Save the modified document.
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);
            doc.Save(outputPath);
            Console.WriteLine($"Converted '{fileName}' and saved to '{outputPath}'.");
        }

        Console.WriteLine("Processing complete.");
    }

    // Helper method to create a simple document with a numbered list if it doesn't exist.
    private static void CreateSampleDocumentIfMissing(string path, string title)
    {
        if (File.Exists(path))
            return;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln(title);
        builder.Writeln();

        // Create a numbered list.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        for (int i = 1; i <= 3; i++)
        {
            Paragraph para = builder.InsertParagraph();
            para.ListFormat.List = numberedList;
            para.ListFormat.ListLevelNumber = 0;
            builder.Writeln($"Item {i}");
        }

        doc.Save(path);
    }
}
