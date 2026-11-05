using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the font for the text that will be added.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;

        // Add a paragraph with sample text.
        builder.Writeln("Hello, world with embedded font!");

        // Validate that the font properties were applied.
        if (builder.Font.Name != "Arial" || builder.Font.Size != 24)
        {
            throw new InvalidOperationException("Font properties were not set correctly.");
        }

        // Configure PDF save options to embed all used fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Save the document as PDF.
        const string outputPath = "EmbeddedFonts.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!System.IO.File.Exists(outputPath))
        {
            throw new System.IO.FileNotFoundException("The PDF file was not created.", outputPath);
        }
    }
}
