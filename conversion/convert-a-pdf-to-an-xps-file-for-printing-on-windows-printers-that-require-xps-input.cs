using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document and save it as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content generated for XPS conversion.");
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the generated PDF.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to XPS using XpsSaveOptions.
        const string xpsPath = "output.xps";
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        pdfDoc.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created as expected.");

        // Optional: clean up intermediate PDF if not needed.
        // File.Delete(pdfPath);
    }
}
