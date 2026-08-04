using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class ListNestingExample
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list (default template has up to 9 levels).
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Add items with increasing list levels, including levels beyond the supported 9.
        // List levels are zero‑based (0..8). Levels 9 and above should fall back to plain text.
        for (int i = 0; i <= 10; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // Set desired level.
            builder.Writeln($"Item at level {i}");
        }

        // End the list formatting.
        builder.ListFormat.List = null;

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "ListNesting.docx");
        doc.Save(docPath);

        // Reload the document to verify the applied formatting.
        Document loadedDoc = new Document(docPath);
        List<bool> isListItemFlags = new List<bool>();

        // Examine each paragraph that we added.
        foreach (Paragraph para in loadedDoc.FirstSection.Body.Paragraphs)
        {
            // Skip empty paragraphs that may exist before or after our list.
            string text = para.GetText().Trim();
            if (string.IsNullOrEmpty(text) || !text.StartsWith("Item at level"))
                continue;

            // Determine whether the paragraph is recognized as a list item.
            bool isListItem = para.ListFormat.IsListItem;
            isListItemFlags.Add(isListItem);
        }

        // Output verification results.
        Console.WriteLine("Verification of list nesting levels (true = list item, false = plain text):");
        for (int i = 0; i < isListItemFlags.Count; i++)
        {
            Console.WriteLine($"Level {i}: {isListItemFlags[i]}");
        }

        // Expected: levels 0‑8 are true, levels 9‑10 are false.
    }
}
