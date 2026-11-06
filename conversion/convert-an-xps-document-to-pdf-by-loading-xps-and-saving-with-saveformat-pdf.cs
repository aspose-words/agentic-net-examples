using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample XPS document created for conversion to PDF.");

        // Save the document as XPS (optional intermediate file).
        string xpsPath = "sample.xps";
        sourceDoc.Save(xpsPath, SaveFormat.Xps);

        // Convert the original document directly to PDF.
        string pdfPath = "output.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The PDF file was not created.");
        }

        // Clean up the intermediate XPS file.
        if (File.Exists(xpsPath))
        {
            File.Delete(xpsPath);
        }

        // Success – no interactive output required.
    }
}
