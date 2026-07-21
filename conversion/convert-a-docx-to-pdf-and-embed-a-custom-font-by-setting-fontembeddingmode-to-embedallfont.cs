using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document with the Arial font.");
        const string inputPath = "sample.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document doc = new Document(inputPath);

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save as PDF with embedded fonts.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: clean up temporary files.
        // File.Delete(inputPath);
    }
}
