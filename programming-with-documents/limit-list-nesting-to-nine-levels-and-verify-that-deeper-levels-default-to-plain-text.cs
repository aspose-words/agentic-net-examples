using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "ListNesting.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list (default template has 9 levels).
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add items for the nine supported levels (0‑8).
        for (int level = 0; level < 9; level++)
        {
            builder.ListFormat.ListLevelNumber = level; // 0‑based level index.
            builder.Writeln($"Level {level + 1}");
        }

        // Attempt to set a level beyond the supported range (level 9 = 10th level).
        builder.ListFormat.ListLevelNumber = 9; // Exceeds the maximum of 8.
        builder.Writeln("Level 10 (should be plain text)");

        // Retrieve the last paragraph that was just added.
        Paragraph lastParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true).Last();

        // Verify that the paragraph is not treated as a list item.
        bool isListItem = lastParagraph.ListFormat.IsListItem;

        // Verify that its outline level defaults to BodyText.
        OutlineLevel outlineLevel = lastParagraph.ParagraphFormat.OutlineLevel;

        // Output verification results.
        Console.WriteLine($"Paragraph at level 10 is list item: {isListItem}");
        Console.WriteLine($"Outline level of paragraph at level 10: {outlineLevel}");

        // Save the document.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
