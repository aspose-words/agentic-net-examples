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

        // Set the default font size (in points) for all text that will be inserted afterwards.
        builder.Font.Size = 14;

        // Simple validation to ensure the size was applied.
        if (builder.Font.Size != 14)
            throw new InvalidOperationException("Failed to set the default font size.");

        // Insert paragraphs; they inherit the default font size set above.
        builder.Writeln("First paragraph with the default font size.");
        builder.Writeln("Second paragraph with the same default font size.");
        builder.Writeln("Third paragraph also using the default font size.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DefaultFontSize.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The document was not saved.", outputPath);
    }
}
