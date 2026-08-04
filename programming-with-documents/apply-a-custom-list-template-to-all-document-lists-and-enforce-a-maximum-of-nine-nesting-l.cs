using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a custom list template (uppercase letters with a dot, e.g., A., B., C.).
        List customList = doc.Lists.Add(ListTemplate.NumberUppercaseLetterDot);

        // Apply the custom list to the builder.
        builder.ListFormat.List = customList;

        // Add list items with increasing nesting levels.
        // Word supports up to 9 levels (0‑8). Levels beyond 8 are clamped to 8.
        for (int i = 0; i < 12; i++)
        {
            int level = Math.Min(i, 8); // Enforce maximum of nine nesting levels.
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Item at original level {i}, applied level {level}");
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomListTemplate.docx");
        doc.Save(outputPath);
    }
}
