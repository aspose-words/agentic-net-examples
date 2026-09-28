using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a Run with sample text.
        Run run = new Run(doc, "Sample text for font size change.");

        // Change the font size of the Run to 14 points.
        run.Font.Size = 14;

        // Validate that the font size was set correctly.
        if (run.Font.Size != 14)
        {
            Console.WriteLine("Font size was not set correctly.");
        }

        // Insert the Run into the document.
        builder.InsertNode(run);

        // Define output file path.
        string outputPath = "Output.docx";

        // Save the document to disk.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
