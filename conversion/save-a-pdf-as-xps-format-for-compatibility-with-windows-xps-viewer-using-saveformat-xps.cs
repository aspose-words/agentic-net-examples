using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content generated for conversion to XPS.");
        sourceDoc.Save("input.pdf", SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document("input.pdf");

        // Convert and save the document as XPS.
        pdfDoc.Save("output.xps", SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists("output.xps"))
            throw new InvalidOperationException("Expected output XPS was not created.");
    }
}
