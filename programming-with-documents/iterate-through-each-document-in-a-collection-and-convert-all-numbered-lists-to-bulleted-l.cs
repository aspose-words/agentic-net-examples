using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for the output documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a collection of sample documents.
        List<Document> documents = new List<Document>
        {
            CreateNumberedDocument("First document with numbered list."),
            CreateNumberedDocument("Second document with another numbered list.")
        };

        // Process each document: convert numbered lists to bulleted lists.
        int docIndex = 1;
        foreach (Document doc in documents)
        {
            // Ensure the document has at least one section.
            doc.EnsureMinimum();

            // Create a single bulleted list that will replace all numbered lists in this document.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Iterate over all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph paragraph in paragraphs)
            {
                // Check if the paragraph is part of a list.
                if (paragraph.ListFormat.IsListItem)
                {
                    // Determine whether the current list is numbered.
                    List currentList = paragraph.ListFormat.List;
                    int levelIndex = paragraph.ListFormat.ListLevelNumber;
                    if (currentList != null &&
                        currentList.ListLevels[levelIndex].NumberStyle != NumberStyle.Bullet)
                    {
                        // Replace the numbered list with the bulleted list, preserving the level.
                        paragraph.ListFormat.List = bulletList;
                        paragraph.ListFormat.ListLevelNumber = levelIndex;
                    }
                }
            }

            // Save the modified document.
            string outPath = Path.Combine(outputDir, $"ModifiedDocument{docIndex}.docx");
            doc.Save(outPath);
            docIndex++;
        }
    }

    // Helper method to create a sample document that contains a numbered list.
    private static Document CreateNumberedDocument(string title)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln(title);
        builder.Writeln();

        // Create a numbered list using the default numbered template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Add several items to the numbered list.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Numbered item {i}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        return doc;
    }
}
