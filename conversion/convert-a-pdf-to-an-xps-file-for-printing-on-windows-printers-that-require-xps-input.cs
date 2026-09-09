using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample content for PDF to XPS conversion.");

        // Save the document as a PDF file (the source for conversion).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the previously saved PDF.
        Document pdfDocument = new Document(pdfPath);

        // Convert the PDF to XPS using XpsSaveOptions.
        const string xpsPath = "output.xps";
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        pdfDocument.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Clean up temporary PDF if desired (optional).
        // File.Delete(pdfPath);
    }
}
