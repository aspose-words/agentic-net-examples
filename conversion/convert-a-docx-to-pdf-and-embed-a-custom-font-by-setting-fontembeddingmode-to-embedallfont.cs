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
        builder.Font.Name = "Courier New"; // Use a non‑standard font to demonstrate embedding.
        builder.Writeln("Sample text with a custom font.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document("input.docx");

        // Set PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Convert and save as PDF.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
