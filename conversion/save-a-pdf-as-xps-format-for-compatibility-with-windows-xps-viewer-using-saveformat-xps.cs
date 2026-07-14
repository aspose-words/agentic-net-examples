using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string xpsPath = "sample.xps";

        // 1. Create a sample Word document and add some content.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("It will be saved as PDF and then converted to XPS.");

        // 2. Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // 3. Load the PDF file.
        Document pdfDoc = new Document(pdfPath);

        // 4. Save the loaded PDF as XPS using SaveFormat.Xps.
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // 5. Verify that the XPS file exists.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("XPS file was not created.");

        // Optional: Inform the user that conversion succeeded.
        Console.WriteLine("PDF successfully converted to XPS.");
    }
}
