using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document using Arial font.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Configure PDF save options to embed all fonts.
        // The enum value for embedding all fonts is the first value (0) in PdfFontEmbeddingMode.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = (PdfFontEmbeddingMode)0 // EmbedAllFonts
        };

        // Save as PDF with the specified options.
        doc.Save("output.pdf", pdfOptions);

        // Validate that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
