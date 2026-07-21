using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("First page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the PDF.");

        const string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF while ignoring errors (recover mode).
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX, preserving page breaks.
        const string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX output file was not created.");

        // Clean up temporary files (optional).
        File.Delete(pdfPath);
    }
}
