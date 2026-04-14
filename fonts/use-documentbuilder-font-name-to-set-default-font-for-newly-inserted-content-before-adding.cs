using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font name for all subsequently inserted text.
        builder.Font.Name = "Arial";

        // Validate that the font name was set correctly.
        if (builder.Font.Name != "Arial")
            throw new InvalidOperationException("Failed to set the default font name.");

        // Insert some text using the default font.
        builder.Writeln("This text uses the default Arial font.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
