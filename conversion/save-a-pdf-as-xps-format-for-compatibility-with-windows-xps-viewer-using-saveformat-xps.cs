using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample content for PDF conversion to XPS.");
        string pdfPath = "sample.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Save the loaded PDF as XPS.
        string xpsPath = "output.xps";
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("Expected XPS output was not created.");

        // Optional cleanup (comment out if you want to keep the files).
        // File.Delete(pdfPath);
        // File.Delete(xpsPath);
    }
}
