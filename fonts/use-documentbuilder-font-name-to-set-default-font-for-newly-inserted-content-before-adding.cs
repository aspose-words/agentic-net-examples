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
        string defaultFont = "Arial";
        builder.Font.Name = defaultFont;

        // Validate that the font name was set correctly.
        if (builder.Font.Name != defaultFont)
            throw new InvalidOperationException("Failed to set the default font.");

        // Insert text using the default font.
        builder.Writeln("This text uses the default font set via DocumentBuilder.Font.Name.");

        // Change the font for a new paragraph to demonstrate that the previous setting was the default.
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses a different font.");

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontExample.docx");
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
