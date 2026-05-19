using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Folder for output documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Prepare a collection of sample documents.
        List<Document> documents = new List<Document>();

        // Document 1 – simple numbered list.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        List numberedList1 = doc1.Lists.Add(ListTemplate.NumberDefault);
        builder1.ListFormat.List = numberedList1;
        builder1.Writeln("Doc1 – Item 1");
        builder1.Writeln("Doc1 – Item 2");
        builder1.Writeln("Doc1 – Item 3");
        builder1.ListFormat.RemoveNumbers(); // End the list.
        documents.Add(doc1);
        doc1.Save(Path.Combine(outputDir, "Original_1.docx"));

        // Document 2 – numbered list with two levels.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        List numberedList2 = doc2.Lists.Add(ListTemplate.NumberDefault);
        builder2.ListFormat.List = numberedList2;
        builder2.Writeln("Doc2 – Level 0 – Item A");
        builder2.ListFormat.ListLevelNumber = 1; // Indent to level 1.
        builder2.Writeln("Doc2 – Level 1 – Subitem A1");
        builder2.Writeln("Doc2 – Level 1 – Subitem A2");
        builder2.ListFormat.ListLevelNumber = 0; // Back to level 0.
        builder2.Writeln("Doc2 – Level 0 – Item B");
        builder2.ListFormat.RemoveNumbers();
        documents.Add(doc2);
        doc2.Save(Path.Combine(outputDir, "Original_2.docx"));

        // Process each document: convert numbered lists to bulleted lists.
        for (int i = 0; i < documents.Count; i++)
        {
            Document doc = documents[i];
            ConvertNumberedListsToBullets(doc);
            string outPath = Path.Combine(outputDir, $"Converted_{i + 1}.docx");
            doc.Save(outPath);
        }
    }

    // Converts every numbered list in the given document to a bulleted list.
    private static void ConvertNumberedListsToBullets(Document doc)
    {
        // Create a bulleted list that will replace numbered lists.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        foreach (Paragraph para in paragraphs.OfType<Paragraph>())
        {
            // Process only paragraphs that are part of a list.
            if (!para.ListFormat.IsListItem)
                continue;

            // Identify the list and level currently applied to the paragraph.
            List currentList = para.ListFormat.List;
            int level = para.ListFormat.ListLevelNumber;

            // Guard against unexpected level values.
            if (currentList == null || level < 0 || level >= currentList.ListLevels.Count)
                continue;

            // Determine if the current list uses a numbered style.
            NumberStyle style = currentList.ListLevels[level].NumberStyle;
            if (style == NumberStyle.Bullet)
                continue; // Already a bulleted list.

            // Apply the bulleted list while preserving the original level.
            para.ListFormat.List = bulletList;
            para.ListFormat.ListLevelNumber = level;
        }
    }
}
