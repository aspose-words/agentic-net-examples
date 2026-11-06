using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("List with varying nesting levels (max 9 levels supported by Word):");

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Insert items with increasing list levels (0‑based). Word supports levels 0‑8 (9 levels).
        const int totalItems = 12; // Create more items than the supported levels.
        for (int i = 0; i < totalItems; i++)
        {
            // Set the desired list level. Levels beyond 8 will be treated as plain text.
            builder.ListFormat.ListLevelNumber = i; // 0‑based level.
            builder.Writeln($"Item at level {i + 1}");
        }

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        const string fileName = "NestedList.docx";
        doc.Save(fileName);

        // Reload the document to verify the behavior.
        Document loaded = new Document(fileName);
        bool verificationPassed = true;
        const int maxSupportedLevel = 9; // Word supports up to 9 levels (0‑8).

        foreach (Paragraph para in loaded.GetChildNodes(NodeType.Paragraph, true))
        {
            string text = para.GetText().Trim();

            // Skip empty paragraphs.
            if (string.IsNullOrEmpty(text))
                continue;

            // Identify paragraphs that were intended to be list items.
            if (text.StartsWith("Item at level"))
            {
                // Extract the level number from the paragraph text.
                // Expected format: "Item at level X"
                string[] parts = text.Split(' ');
                if (parts.Length >= 4 && int.TryParse(parts[3], out int levelFromText))
                {
                    // ListFormat.IsListItem indicates whether the paragraph is a list item.
                    bool isListItem = para.ListFormat.IsListItem;
                    int listLevel = para.ListFormat.ListLevelNumber; // 0‑based.

                    if (levelFromText <= maxSupportedLevel)
                    {
                        // Levels within the supported range should be list items.
                        if (!isListItem || listLevel != levelFromText - 1)
                            verificationPassed = false;
                    }
                    else
                    {
                        // Levels beyond the supported range should be plain text (not a list item).
                        if (isListItem)
                            verificationPassed = false;
                    }
                }
                else
                {
                    verificationPassed = false;
                }
            }
        }

        Console.WriteLine(verificationPassed ? "Verification passed." : "Verification failed.");
    }
}
