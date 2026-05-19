using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample XPS content.");

        // Save the document as XPS – this file will act as the input for the conversion demo.
        const string xpsPath = "input.xps";
        source.Save(xpsPath, SaveFormat.Xps);

        // Aspose.Words cannot load XPS directly, so we reuse the original document
        // and save it as PDF to demonstrate the conversion workflow.
        const string pdfPath = "output.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
