using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add items with increasing list levels (0‑based). Aspose.Words supports up to 9 levels (0‑8).
        // We add 11 items to exceed the limit.
        for (int i = 0; i < 11; i++)
        {
            // Set the desired list level.
            builder.ListFormat.ListLevelNumber = i;
            builder.Writeln($"Item at level {i + 1}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        const string fileName = "ListNesting.docx";
        doc.Save(fileName);

        // Reload the document to verify the list nesting behavior.
        Document loadedDoc = new Document(fileName);
        bool verificationPassed = true;

        // Iterate through all paragraphs and check list status.
        foreach (Paragraph para in loadedDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            bool isListItem = para.ListFormat.IsListItem;
            int level = para.ListFormat.ListLevelNumber; // 0‑based level; -1 if not a list item.

            // Levels 0‑8 should be list items; deeper levels should default to plain text.
            if (level >= 0 && level <= 8)
            {
                if (!isListItem)
                {
                    verificationPassed = false;
                    Console.WriteLine($"Paragraph \"{para.GetText().Trim()}\" expected to be a list item at level {level + 1} but is not.");
                }
            }
            else if (level > 8)
            {
                if (isListItem)
                {
                    verificationPassed = false;
                    Console.WriteLine($"Paragraph \"{para.GetText().Trim()}\" exceeds max nesting level but is still a list item.");
                }
            }
        }

        // Output verification result.
        Console.WriteLine(verificationPassed
            ? "Verification passed: nesting limited to nine levels; deeper levels are plain text."
            : "Verification failed: list nesting behavior is not as expected.");
    }
}
