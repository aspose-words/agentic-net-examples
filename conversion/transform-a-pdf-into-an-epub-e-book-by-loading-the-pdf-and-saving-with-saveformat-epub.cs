using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content for EPUB conversion.");
        source.Save("sample.pdf", SaveFormat.Pdf);

        // Load the PDF and convert it to EPUB.
        Document pdfDoc = new Document("sample.pdf");
        pdfDoc.Save("output.epub", SaveFormat.Epub);

        // Verify that the EPUB file was created.
        if (!File.Exists("output.epub"))
            throw new InvalidOperationException("The EPUB file was not created.");
    }
}
