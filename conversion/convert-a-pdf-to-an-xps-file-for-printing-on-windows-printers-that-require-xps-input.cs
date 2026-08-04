using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string pdfPath = "sample.pdf";
        string xpsPath = "output.xps";

        // Create a simple document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document created for XPS conversion.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to XPS using XpsSaveOptions.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        pdfDoc.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optionally, clean up the intermediate PDF file.
        // File.Delete(pdfPath);
    }
}
