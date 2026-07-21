using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and the resulting XPS file.
        const string pdfPath = "sample.pdf";
        const string xpsPath = "sample.xps";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF (input source).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF that will be converted to XPS.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // ---------------------------------------------------------------
        // 2. Load the generated PDF document.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // ---------------------------------------------------------------
        // 3. Convert the PDF to XPS using XpsSaveOptions.
        // ---------------------------------------------------------------
        XpsSaveOptions xpsOptions = new XpsSaveOptions(); // default options
        pdfDoc.Save(xpsPath, xpsOptions);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");
    }
}
