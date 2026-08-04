using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple PDF document as the source.
        Document sourcePdf = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourcePdf);
        builder.Writeln("This is a sample PDF that will be converted to EPUB.");
        const string pdfPath = "input.pdf";
        sourcePdf.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert and save the document as EPUB.
        const string epubPath = "output.epub";
        pdfDoc.Save(epubPath, SaveFormat.Epub);

        // Verify that the EPUB file was created.
        if (!File.Exists(epubPath))
            throw new InvalidOperationException("The EPUB file was not created.");

        // Optional: indicate success (no console interaction required).
        // The program will exit automatically.
    }
}
