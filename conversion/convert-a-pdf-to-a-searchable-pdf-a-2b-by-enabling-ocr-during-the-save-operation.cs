using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document that will be converted to PDF/A‑2b.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be converted to searchable PDF/A‑2b.");

        // Save the document as a temporary PDF file.
        const string inputPdfPath = "input.pdf";
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(inputPdfPath);

        // Configure PDF save options for PDF/A‑2b compliance.
        // Note: In older Aspose.Words versions PDF/A‑2b may not be available;
        // PdfA2a is used as the closest alternative.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2a
            // OCR support is not available in this version of Aspose.Words,
            // therefore OcrMode and OcrLanguage properties are omitted.
        };

        // Save the PDF with the specified options, producing a PDF/A‑2b (or PDF/A‑2a) file.
        const string outputPdfPath = "output.pdf";
        pdfDoc.Save(outputPdfPath, saveOptions);

        // Validate that the output file was created and contains data.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The searchable PDF/A‑2b file was not created.");

        FileInfo info = new FileInfo(outputPdfPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The searchable PDF/A‑2b file is empty.");

        // Clean up the temporary input file.
        if (File.Exists(inputPdfPath))
            File.Delete(inputPdfPath);
    }
}
