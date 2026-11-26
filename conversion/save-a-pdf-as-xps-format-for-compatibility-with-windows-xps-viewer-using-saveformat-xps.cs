using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for PDF to XPS conversion.");

        // Define file names.
        string pdfPath = "sample.pdf";
        string xpsPath = "sample.xps";

        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected PDF file was not created.");

        // Save the same document as XPS.
        doc.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("Expected XPS file was not created.");

        // Optionally, clean up the sample files (comment out if you want to keep them).
        // File.Delete(pdfPath);
        // File.Delete(xpsPath);
    }
}
