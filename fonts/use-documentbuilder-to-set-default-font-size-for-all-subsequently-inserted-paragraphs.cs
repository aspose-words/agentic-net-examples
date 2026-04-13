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
        double defaultFontSize = 14.0;
        builder.Font.Size = defaultFontSize;

        // Validate that the font size was set correctly.
        if (builder.Font.Size != defaultFontSize)
            throw new InvalidOperationException("Failed to set the default font size.");

        // Insert paragraphs; they inherit the default font size.
        builder.Writeln("First paragraph with the default font size.");
        builder.Writeln("Second paragraph with the default font size.");

        // Save the document.
        string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontSize.docx");
        doc.Save(outputFile);

        // Ensure the file was created.
        if (!File.Exists(outputFile))
            throw new FileNotFoundException("The document was not saved correctly.", outputFile);
    }
}
