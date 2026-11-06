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

        // Create a numbered list using a built‑in list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list; // Apply the list to the builder.

        // Attempt to write list items with increasing nesting levels.
        // Aspose.Words supports a maximum of 9 levels (0‑8). Level 9 will throw.
        for (int level = 0; level <= 9; level++)
        {
            try
            {
                // Set the current list level.
                builder.ListFormat.ListLevelNumber = level;

                // Write a paragraph for this level.
                builder.Writeln($"List item at level {level}");
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Catch the exception that occurs when the level exceeds the allowed range.
                Console.WriteLine($"Caught exception for level {level}: {ex.Message}");
                break; // Exit the loop after handling the error.
            }
        }

        // Remove list formatting to return the builder to normal paragraph mode.
        builder.ListFormat.RemoveNumbers();

        // Save the document (optional, demonstrates that the file is created).
        doc.Save("Output.docx");
    }
}
