using System;
using System.IO;
using Aspose.Words;

public class PdfToXpsConverter
{
    public static void Main()
    {
        // Create a sample document with some content.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document that will be converted to XPS.");

        // Save the sample document as PDF (input for conversion).
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert and save the PDF as XPS.
        string xpsPath = "output.xps";
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // Validate that the XPS file was created.
        if (!File.Exists(xpsPath))
        {
            throw new InvalidOperationException("Expected output XPS file was not created.");
        }

        // Optionally, clean up the temporary PDF file.
        // File.Delete(pdfPath);
    }
}
