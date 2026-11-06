using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some content.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document that will be saved as PDF and then converted to XPS.");

        // Define file names for the intermediate PDF and final XPS.
        string pdfPath = "sample.pdf";
        string xpsPath = "output.xps";

        // Save the document as PDF (this will serve as the input PDF for conversion).
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The intermediate PDF file was not created.");

        // Load the PDF file.
        Document pdfDoc = new Document(pdfPath);

        // Convert the loaded PDF to XPS.
        pdfDoc.Save(xpsPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(xpsPath))
            throw new InvalidOperationException("The XPS file was not created.");

        // Optionally, clean up the intermediate PDF file.
        try
        {
            File.Delete(pdfPath);
        }
        catch
        {
            // Ignore any errors during cleanup.
        }

        // Indicate successful conversion (no console input required).
        Console.WriteLine("PDF successfully converted to XPS.");
    }
}
