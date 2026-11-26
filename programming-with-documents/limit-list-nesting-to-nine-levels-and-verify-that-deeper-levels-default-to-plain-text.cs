using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "ListNesting.docx";

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a default numbered list to start list formatting.
        builder.ListFormat.ApplyNumberDefault();

        // Attempt to create list items for 11 levels (0‑based indices 0..10).
        // Aspose.Words supports only 9 levels (0..8). Levels beyond that should fall back to plain text.
        for (int level = 0; level < 11; level++)
        {
            try
            {
                // Set the desired list level. If the level is out of range, Aspose.Words will ignore the list formatting.
                builder.ListFormat.ListLevelNumber = level;
            }
            catch (ArgumentOutOfRangeException)
            {
                // If an exception occurs, reset to plain text formatting.
                builder.ListFormat.RemoveNumbers();
            }

            // Write the paragraph text.
            builder.Writeln($"Level {level + 1}");
        }

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save(outputPath);

        // Reload the document to verify list nesting behavior.
        Document loadedDoc = new Document(outputPath);
        ParagraphCollection paragraphs = loadedDoc.FirstSection.Body.Paragraphs;

        // Verify each paragraph.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = paragraphs[i];
            bool isListItem = para.ListFormat.IsListItem;
            int listLevel = para.ListFormat.ListLevelNumber; // Returns 0‑based level; -1 if not a list item.

            // Expected: first 9 paragraphs are list items, deeper ones are plain text.
            bool shouldBeList = i < 9; // Zero‑based index; first 9 paragraphs correspond to levels 0‑8.

            if (isListItem != shouldBeList)
            {
                Console.WriteLine($"Verification failed at paragraph {i + 1}: " +
                                  $"IsListItem={isListItem}, Expected={shouldBeList}");
            }
            else
            {
                Console.WriteLine($"Paragraph {i + 1} verified. " +
                                  $"IsListItem={isListItem}, Level={(listLevel >= 0 ? listLevel + 1 : 0)}");
            }
        }

        // Clean up: optionally delete the file after verification.
        // File.Delete(outputPath);
    }
}
