using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font size for all subsequently inserted text.
        double defaultFontSize = 14.0;
        builder.Font.Size = defaultFontSize;

        // Validate that the font size was applied.
        if (builder.Font.Size != defaultFontSize)
            throw new InvalidOperationException("Failed to set the default font size.");

        // Insert paragraphs; they will use the default font size set above.
        builder.Writeln("First paragraph with the default font size.");
        builder.Writeln("Second paragraph with the same default font size.");

        // Save the document to a file.
        string outputFile = "DefaultFontSize.docx";
        doc.Save(outputFile);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputFile))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
