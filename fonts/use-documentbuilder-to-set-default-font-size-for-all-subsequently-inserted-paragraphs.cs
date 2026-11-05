using System;
using System.IO;
using Aspose.Words;
using Aspose.Drawing; // Required package, not used directly
using Newtonsoft.Json; // Required package, not used directly

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font size for all subsequently inserted paragraphs.
        builder.Font.Size = 14; // 14 points

        // Validate that the font size was set correctly.
        if (builder.Font.Size != 14)
        {
            throw new InvalidOperationException("Failed to set the default font size.");
        }

        // Insert paragraphs that will use the default font size.
        builder.Writeln("First paragraph using the default font size.");
        builder.Writeln("Second paragraph using the default font size.");

        // Save the document to disk.
        string outputPath = "DefaultFontSize.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The document was not saved correctly.", outputPath);
        }
    }
}
