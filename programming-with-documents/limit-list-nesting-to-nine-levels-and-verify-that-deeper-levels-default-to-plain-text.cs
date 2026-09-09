using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Define a folder for output files and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "ListNesting.docx");

        // Create a new blank document and a DocumentBuilder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list that supports up to 9 levels (0‑8).
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add items for the nine supported levels.
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // 0‑based level index.
            builder.Writeln($"Level {i + 1}");
        }

        // Attempt to set a level beyond the supported range (level 9, i.e., the 10th level).
        // Aspose.Words will treat this paragraph as plain text, not as a list item.
        builder.ListFormat.ListLevelNumber = 9; // Exceeds the maximum of 8.
        builder.Writeln("Level 10 (should be plain text)");

        // Retrieve the last paragraph to verify its list status.
        Paragraph lastParagraph = doc.FirstSection.Body.Paragraphs[doc.FirstSection.Body.Paragraphs.Count - 1];
        bool isListItem = lastParagraph.ListFormat.IsListItem;

        // Output verification result.
        Console.WriteLine($"Paragraph at level 10 is a list item: {isListItem}");

        // Save the document.
        doc.Save(outputPath);
    }
}
