using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font name for all subsequently inserted text.
        builder.Font.Name = "Arial";

        // Verify that the font name was set correctly.
        if (builder.Font.Name != "Arial")
        {
            Console.WriteLine("Failed to set the default font.");
            return;
        }

        // Insert text that will use the default font.
        builder.Writeln("This text is formatted with the default Arial font.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontExample.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Confirm that the file was created.
        Console.WriteLine(File.Exists(outputPath) ? "Document saved successfully." : "Document save failed.");
    }
}
