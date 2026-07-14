using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the second page.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF with options that ignore errors during loading.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Attempt to recover from any errors encountered while loading.
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX, preserving page breaks.
        const string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Clean up temporary PDF file (optional).
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
