using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document and save it as PDF (input for conversion).
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for PDF to XPS conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDocument = new Document(pdfPath);

        // Convert and save the PDF as XPS.
        string xpsPath = "output.xps";
        pdfDocument.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
        {
            throw new InvalidOperationException("Expected XPS output was not created.");
        }

        // Clean up intermediate PDF if desired (optional).
        // File.Delete(pdfPath);
    }
}
