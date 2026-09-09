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
        // Step 1: Create a sample DOCX document with a non‑standard font.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        // "Courier New" is not a standard Windows font for PDF embedding,
        // so it will be treated as a custom font.
        builder.Font.Name = "Courier New";
        builder.Writeln("This paragraph uses the Courier New font and will be embedded in the PDF.");

        // Save the DOCX to disk (input bootstrap rule).
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX file that we just created.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Step 3: Configure PDF save options to embed all fonts.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed every font used in the document.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            // Optional: embed the full font data (no subsetting).
            EmbedFullFonts = true
        };

        // -----------------------------------------------------------------
        // Step 4: Save the document as PDF using the configured options.
        // -----------------------------------------------------------------
        doc.Save(outputPath, pdfOptions);

        // -----------------------------------------------------------------
        // Step 5: Validate that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The expected PDF output file was not created.");
        }

        // Optionally, report success (no interactive input required).
        Console.WriteLine("DOCX successfully converted to PDF with all fonts embedded.");
    }
}
