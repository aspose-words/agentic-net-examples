using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate PDF and final EPUB files
        const string pdfPath = "sample.pdf";
        const string epubPath = "sample.epub";

        // Create a simple Word document with some text
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF that will be converted to EPUB.");

        // Save the document as PDF
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF document
        Document pdfDoc = new Document(pdfPath);

        // Convert the loaded PDF to EPUB
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The EPUB file was not created.");
    }
}
