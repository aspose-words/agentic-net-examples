using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define temporary file paths.
        const string pdfPath = "sample.pdf";
        const string docxPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF with options that ignore errors.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Explicitly set the recovery mode to try to recover the document if it contains errors.
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };

        Document pdfDoc = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Save the loaded document as DOCX while preserving page breaks.
        // -----------------------------------------------------------------
        // No special save options are required; Aspose.Words preserves page breaks by default.
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Verify that the DOCX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Clean up the temporary PDF file (optional).
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
