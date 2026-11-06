using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a font that is likely to be available.
        string fontName = "Arial";
        builder.Font.Name = fontName;

        // Validate that the font was set correctly.
        if (builder.Font.Name != fontName)
        {
            throw new InvalidOperationException("Failed to set the font name.");
        }

        // Add some text to the document.
        builder.Writeln("This text should be embedded in the PDF to preserve its appearance.");

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true
        };

        // Define the output PDF file path.
        string outputPath = "EmbeddedFontOutput.pdf";

        // Save the document as PDF with font embedding enabled.
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The PDF file was not created.", outputPath);
        }
    }
}
