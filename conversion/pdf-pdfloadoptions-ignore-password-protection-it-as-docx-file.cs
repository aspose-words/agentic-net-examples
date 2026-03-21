using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Create temporary file paths.
        string tempDir = Path.GetTempPath();
        string pdfPath = Path.Combine(tempDir, "source.pdf");
        string docxPath = Path.Combine(tempDir, "converted.docx");

        // Create a simple Word document and save it as PDF to ensure the source PDF exists.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.Writeln("This is a sample PDF generated for conversion.");
        tempDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF with PdfLoadOptions (no password provided, so any protection is ignored).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        doc.Save(docxPath, SaveFormat.Docx);

        Console.WriteLine($"PDF converted to DOCX successfully.");
        Console.WriteLine($"Source PDF: {pdfPath}");
        Console.WriteLine($"Output DOCX: {docxPath}");
    }
}
