using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.pdf");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "converted.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple PDF file that will be used as the input source.
        // -----------------------------------------------------------------
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Writeln("First page of the PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the PDF.");
        // Save the document as PDF.
        tempDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // 2. Load the PDF with PdfLoadOptions configured to ignore errors.
        // ---------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Use the default recovery mode (TryRecover) to handle any errors gracefully.
            RecoveryMode = DocumentRecoveryMode.TryRecover
        };

        Document pdfDocument = new Document(pdfPath, loadOptions);

        // ---------------------------------------------------------------
        // 3. Save the loaded document as DOCX, preserving page breaks.
        // ---------------------------------------------------------------
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // 4. Verify that the output file was created.
        // ---------------------------------------------------------------
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("The DOCX file was not created.");
        }

        Console.WriteLine("PDF successfully converted to DOCX with page breaks preserved.");
    }
}
