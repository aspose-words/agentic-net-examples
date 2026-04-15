using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ListNesting.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add items for the nine supported list levels (0‑8).
        for (int level = 0; level < 9; level++)
        {
            builder.ListFormat.ListLevelNumber = level; // valid levels only.
            builder.Writeln($"Level {level + 1}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Add items that would exceed the maximum list depth.
        // These paragraphs are written as plain text (no list formatting).
        for (int extra = 9; extra < 12; extra++)
        {
            builder.Writeln($"Plain level {extra + 1}");
        }

        // Save the document.
        doc.Save(outputPath);

        // Reload the document to verify the structure.
        Document loadedDoc = new Document(outputPath);
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);

        int listItemCount = 0;
        int plainParagraphCount = 0;

        foreach (Paragraph para in paragraphs)
        {
            if (para.ListFormat.IsListItem)
                listItemCount++;
            else
                plainParagraphCount++;
        }

        // Expected: 9 list items (levels 1‑9) and the remaining paragraphs as plain text.
        Console.WriteLine($"List items: {listItemCount} (expected 9)");
        Console.WriteLine($"Plain paragraphs: {plainParagraphCount} (expected 3)");
    }
}
