using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string docxPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF document with page breaks.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page content.");
        // Save as PDF.
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF with PdfLoadOptions that ignore errors.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Use the default recovery mode (TryRecover) to ignore errors.
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Step 3: Save the loaded document as DOCX, preserving page breaks.
        // -----------------------------------------------------------------
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: ensure the DOCX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX output file was not created.");

        // Optional: clean up sample files (comment out if you want to keep them).
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
