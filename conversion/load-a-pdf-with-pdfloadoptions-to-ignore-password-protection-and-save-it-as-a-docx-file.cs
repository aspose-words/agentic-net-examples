using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content.");
        const string pdfPath = "sample.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF using PdfLoadOptions (no password supplied, so any password protection is ignored).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        const string docxPath = "output.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("The DOCX file was not created.");
        }
    }
}
