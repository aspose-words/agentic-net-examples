using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document that will be saved as PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for EPUB conversion.");

        // Save the document as PDF – this will be the input for the conversion.
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the previously saved PDF.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to EPUB.
        const string epubPath = "output.epub";
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The EPUB file was not created.");

        // Optional cleanup of the intermediate PDF.
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
