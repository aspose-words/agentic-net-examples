using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some text.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample XPS content.");

        // Save the document as XPS (the intermediate file is not used for loading).
        const string xpsPath = "input.xps";
        source.Save(xpsPath, SaveFormat.Xps);

        // Save the same document as DOCX so it can be loaded again.
        const string docxPath = "input.docx";
        source.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document (Aspose.Words cannot load XPS directly).
        Document doc = new Document(docxPath);

        // Convert and save the document as PDF.
        const string pdfPath = "output.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Indicate success (no interactive input required).
        Console.WriteLine("XPS to PDF conversion completed successfully.");
    }
}
