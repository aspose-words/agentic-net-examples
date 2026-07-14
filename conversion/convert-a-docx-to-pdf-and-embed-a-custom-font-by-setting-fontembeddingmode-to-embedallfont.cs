using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string docxPath = "sample.docx";
        const string pdfPath = "sample.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";               // Use a common font.
        builder.Writeln("This is a sample document.");
        builder.Font.Name = "Courier New";         // Add a second font.
        builder.Writeln("Text with a different font.");

        // Save the document as DOCX (input for conversion).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure PDF save options to embed all fonts.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF with the specified options.
        loadedDoc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
        {
            throw new InvalidOperationException("PDF conversion failed: output file was not created.");
        }

        // Optional: inform the user (no interactive input required).
        Console.WriteLine($"DOCX successfully converted to PDF with embedded fonts: {pdfPath}");
    }
}
