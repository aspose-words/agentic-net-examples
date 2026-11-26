using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF file using Aspose.Words.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document generated for conversion to XPS.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF file was not created.");

        // Load the PDF file.
        Document pdfDocument = new Document(pdfPath);

        // Convert the loaded PDF to XPS.
        string xpsPath = "output.xps";
        pdfDocument.Save(xpsPath, SaveFormat.Xps);

        // Verify the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optionally, clean up the intermediate PDF file.
        // File.Delete(pdfPath);
    }
}
