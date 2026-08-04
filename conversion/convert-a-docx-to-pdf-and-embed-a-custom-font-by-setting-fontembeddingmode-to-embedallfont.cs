using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "sample.docx";
        const string outputPath = "sample.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial";
        builder.Writeln("This paragraph uses the Arial font.");
        builder.Font.Name = "Times New Roman";
        builder.Writeln("This paragraph uses the Times New Roman font.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Configure PDF save options to embed all fonts.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // -----------------------------------------------------------------
        // 4. Save the document as PDF with the specified options.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optional: clean up generated files (comment out if you want to keep them).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
