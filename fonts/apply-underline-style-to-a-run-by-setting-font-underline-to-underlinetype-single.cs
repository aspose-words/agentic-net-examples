using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a single underline to the current font.
        builder.Font.Underline = Underline.Single;

        // Validate that the underline was set correctly.
        if (builder.Font.Underline != Underline.Single)
            throw new InvalidOperationException("Failed to set underline style.");

        // Add some text to demonstrate the underline.
        builder.Writeln("This text is underlined.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UnderlineExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
