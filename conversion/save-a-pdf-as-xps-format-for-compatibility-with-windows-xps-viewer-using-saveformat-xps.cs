using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for PDF to XPS conversion.");

        // Save the document as PDF.
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Save the PDF as XPS.
        string xpsPath = "output.xps";
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optional: clean up intermediate PDF file.
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
