using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a collection of sample documents.
        List<Document> documents = new List<Document>
        {
            CreateSampleDocumentWithNumberedList(),
            CreateSampleDocumentWithNumberedAndBulletedLists()
        };

        // Process each document: convert numbered lists to bulleted lists.
        int docIndex = 1;
        foreach (Document doc in documents)
        {
            // Ensure the document has a bulleted list template to reuse.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Get all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph paragraph in paragraphs)
            {
                // Skip paragraphs that are not part of a list.
                if (!paragraph.ListFormat.IsListItem)
                    continue;

                // Determine the current list level.
                int levelNumber = paragraph.ListFormat.ListLevelNumber;
                List currentList = paragraph.ListFormat.List;
                ListLevel currentLevel = currentList.ListLevels[levelNumber];

                // If the current level uses a numbered style, replace it with the bulleted list.
                if (currentLevel.NumberStyle != NumberStyle.Bullet)
                {
                    paragraph.ListFormat.List = bulletList;
                    // Preserve the original list level number.
                    paragraph.ListFormat.ListLevelNumber = levelNumber;
                }
            }

            // Save the modified document.
            string outPath = Path.Combine(outputDir, $"Document_{docIndex}_Converted.docx");
            doc.Save(outPath);
            docIndex++;
        }
    }

    // Creates a document that contains a simple numbered list.
    private static Document CreateSampleDocumentWithNumberedList()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a default numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("First numbered item");
        builder.Writeln("Second numbered item");
        builder.Writeln("Third numbered item");
        builder.ListFormat.RemoveNumbers();

        return doc;
    }

    // Creates a document that contains both numbered and bulleted lists.
    private static Document CreateSampleDocumentWithNumberedAndBulletedLists()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Numbered list.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        builder.ListFormat.RemoveNumbers();

        // Bulleted list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Bulleted item A");
        builder.Writeln("Bulleted item B");
        builder.ListFormat.RemoveNumbers();

        return doc;
    }
}
