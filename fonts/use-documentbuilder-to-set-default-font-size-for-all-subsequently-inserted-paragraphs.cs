using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font size for all subsequently inserted text.
        double defaultFontSize = 16.0;
        builder.Font.Size = defaultFontSize;

        // Validate that the font size was set correctly.
        if (builder.Font.Size != defaultFontSize)
        {
            throw new InvalidOperationException("Failed to set the default font size.");
        }

        // Insert paragraphs that will use the default font size.
        builder.Writeln("This paragraph uses the default font size.");
        builder.Writeln("This is another paragraph with the same default font size.");

        // Save the document to disk.
        string outputPath = "DefaultFontSize.docx";
        doc.Save(outputPath);

        // Ensure the output file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
