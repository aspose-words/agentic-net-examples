using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for PDF to XPS conversion.");

        // Save the document as PDF (the source for conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the generated PDF.
        Document pdfDocument = new Document(pdfPath);

        // Prepare XPS save options.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the PDF document as XPS.
        const string xpsPath = "output.xps";
        pdfDocument.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optional cleanup of the intermediate PDF.
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
