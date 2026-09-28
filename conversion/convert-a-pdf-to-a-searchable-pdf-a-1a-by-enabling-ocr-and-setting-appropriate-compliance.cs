using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample PDF file that will serve as the input document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample text for OCR conversion.");
        const string inputPath = "input.pdf";
        sourceDoc.Save(inputPath, SaveFormat.Pdf);

        // 2. Load the PDF that we want to convert to a searchable PDF/A‑1a.
        Document pdfDocument = new Document(inputPath);

        // 3. Configure PDF save options:
        //    - Set PDF/A‑1a compliance.
        //    - OCR settings are omitted because the current Aspose.Words version
        //      does not expose OcrSettings on PdfSaveOptions.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // 4. Save the document as a PDF/A‑1a file.
        const string outputPath = "output.pdf";
        pdfDocument.Save(outputPath, saveOptions);

        // 5. Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF/A‑1a file was not created.");
        }

        // 6. Clean up the temporary input file.
        if (File.Exists(inputPath))
        {
            File.Delete(inputPath);
        }
    }
}
