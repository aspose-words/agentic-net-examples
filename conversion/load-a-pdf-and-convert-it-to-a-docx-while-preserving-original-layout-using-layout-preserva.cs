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
        const string pdfPath = "sample.pdf";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF created by Aspose.Words.");
        builder.Writeln("It contains multiple lines to test layout preservation.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF. Layout preservation is handled automatically by the loader.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // Convert the loaded PDF to DOCX.
        const string docxPath = "converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("DOCX conversion failed: output file not found.");
    }
}
