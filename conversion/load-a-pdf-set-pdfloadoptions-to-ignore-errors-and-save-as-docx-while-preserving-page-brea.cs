using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF with two pages.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        const string pdfPath = "sample.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF while ignoring recoverable errors.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Use TryRecover to attempt recovery and ignore recoverable issues.
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX. Page breaks are preserved automatically.
        const string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Clean up temporary files (optional).
        File.Delete(pdfPath);
    }
}
