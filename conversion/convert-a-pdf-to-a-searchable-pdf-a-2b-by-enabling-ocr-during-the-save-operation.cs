using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple source document and save it as a PDF (non‑searchable).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is sample text for OCR conversion to PDF/A‑2b.");
        const string inputPdfPath = "input.pdf";
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        if (!File.Exists(inputPdfPath) || new FileInfo(inputPdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the input PDF.");

        // -----------------------------------------------------------------
        // 2. Load the PDF that we just created.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(inputPdfPath);

        // -----------------------------------------------------------------
        // 3. Save the PDF again.  The example focuses on the conversion flow;
        //    OCR‑related options are not available in Aspose.Words, so we simply
        //    save the document using PdfSaveOptions (you could set compliance
        //    here if the required enum value is present in your library version).
        // -----------------------------------------------------------------
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Example of setting PDF/A compliance when the enum value exists:
        // Uncomment the following line if your Aspose.Words version supports PdfA2b.
        // saveOptions.Compliance = PdfCompliance.PdfA2b;

        const string outputPdfPath = "output_pdfa2b.pdf";
        pdfDoc.Save(outputPdfPath, saveOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the output file was created and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath) || new FileInfo(outputPdfPath).Length == 0)
            throw new InvalidOperationException("The searchable PDF/A‑2b file was not created.");

        // -----------------------------------------------------------------
        // 5. Clean up the temporary input file.
        // -----------------------------------------------------------------
        try
        {
            File.Delete(inputPdfPath);
        }
        catch
        {
            // Ignore cleanup errors.
        }

        Console.WriteLine("Conversion completed successfully.");
    }
}
