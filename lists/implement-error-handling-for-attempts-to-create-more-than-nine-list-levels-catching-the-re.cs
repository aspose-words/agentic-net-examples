using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list using a built‑in template.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Write list items for all valid levels (0 to 8).
        for (int i = 0; i < 9; i++)
        {
            builder.ListFormat.ListLevelNumber = i; // Valid level.
            builder.Writeln($"Level {i}");
        }

        // Attempt to set a level beyond the allowed range and catch the exception.
        try
        {
            builder.ListFormat.ListLevelNumber = 9; // Invalid level – should throw.
            builder.Writeln("This line will not be added.");
        }
        catch (ArgumentOutOfRangeException ex)
        {
            Console.WriteLine($"Caught expected exception: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Caught unexpected exception: {ex.Message}");
        }

        // End list formatting.
        builder.ListFormat.List = null;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ListLevelExceptionDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
