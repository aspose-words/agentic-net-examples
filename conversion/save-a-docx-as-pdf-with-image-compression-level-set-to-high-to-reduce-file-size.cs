using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string docxPath = "sample.docx";
        const string pdfPath = "sample.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document created for PDF conversion.");
        // Save the document as DOCX (bootstrap step for input file).
        sourceDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure PDF save options with high image compression.
        //    - Use JPEG compression for all images.
        //    - Set JPEG quality to a low value (e.g., 10) to achieve strong compression.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 10 // 0 = worst quality, highest compression.
        };

        // -----------------------------------------------------------------
        // 4. Save the document as PDF using the configured options.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created as expected.");

        // Optional: Output the size of the generated PDF for verification.
        FileInfo info = new FileInfo(pdfPath);
        Console.WriteLine($"PDF saved successfully. Size: {info.Length} bytes.");
    }
}
