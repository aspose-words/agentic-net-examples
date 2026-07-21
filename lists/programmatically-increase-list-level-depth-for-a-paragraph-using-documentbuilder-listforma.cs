using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class ListIndentExample
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Write the first item at level 0.
        builder.Writeln("Level 0");

        // Increase the list level inside a loop and write items at deeper levels.
        for (int i = 1; i <= 5; i++)
        {
            // Increase the current list level by one.
            builder.ListFormat.ListIndent();

            // Write a paragraph at the new list level.
            builder.Writeln($"Level {i}");
        }

        // Return to the original level and end the list.
        for (int i = 0; i < 5; i++)
        {
            builder.ListFormat.ListOutdent();
        }
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ListIndentExample.docx");
        doc.Save(outputPath);
    }
}
