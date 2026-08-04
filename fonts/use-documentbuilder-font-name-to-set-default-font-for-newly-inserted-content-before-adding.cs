using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font that will be applied to all newly inserted text.
        builder.Font.Name = "Arial";

        // Insert some text using the default font.
        builder.Writeln("Hello world! This text uses the default Arial font.");

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DefaultFontExample.docx");

        // Save the document to the specified path.
        doc.Save(outputPath);

        // Validate that the file was created successfully.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
