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

        // Set the default font size that will be applied to all subsequently inserted text.
        const double defaultFontSize = 14.0;
        builder.Font.Size = defaultFontSize;

        // Validate that the font size was set correctly.
        if (builder.Font.Size != defaultFontSize)
            throw new InvalidOperationException("Failed to set the default font size.");

        // Insert several paragraphs; they will inherit the default font size.
        builder.Writeln("First paragraph with default font size.");
        builder.Writeln("Second paragraph with default font size.");
        builder.Writeln("Third paragraph with default font size.");

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DefaultFontSizeExample.docx");

        // Save the document to the file system.
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
    }
}
