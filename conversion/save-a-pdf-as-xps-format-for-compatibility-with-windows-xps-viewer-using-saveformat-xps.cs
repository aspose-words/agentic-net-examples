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
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF we just created.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to XPS.
        const string xpsPath = "sample.xps";
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optional cleanup (comment out if you want to keep the files).
        // File.Delete(pdfPath);
        // File.Delete(xpsPath);
    }
}
