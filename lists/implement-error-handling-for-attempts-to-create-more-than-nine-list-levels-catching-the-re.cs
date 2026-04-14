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

        // Start a default numbered list.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Write items for the valid levels (0 to 8).
        for (int level = 0; level <= 8; level++)
        {
            builder.ListFormat.ListLevelNumber = level;
            builder.Writeln($"Valid level {level}");
        }

        // Attempt to set an invalid level (greater than 8) and handle the exception.
        try
        {
            // This should throw because list levels are limited to 0‑8 (nine levels).
            builder.ListFormat.ListLevelNumber = 9;
            builder.Writeln("This line will not be reached.");
        }
        catch (Exception ex)
        {
            // Output the exception message to the console.
            Console.WriteLine("Caught exception while setting an invalid list level:");
            Console.WriteLine(ex.Message);
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Lists_ErrorHandling.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
