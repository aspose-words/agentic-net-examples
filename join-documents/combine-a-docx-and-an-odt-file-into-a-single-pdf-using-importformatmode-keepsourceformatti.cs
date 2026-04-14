using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // File names for the sample documents and the final PDF.
        const string docxPath = "Sample.docx";
        const string odtPath = "Sample.odt";
        const string pdfPath = "Combined.pdf";

        // -----------------------------------------------------------------
        // Create a DOCX source document.
        // -----------------------------------------------------------------
        var docx = new Document();
        var builderDocx = new DocumentBuilder(docx);
        builderDocx.Writeln("This is content from the DOCX source document.");
        // Save the DOCX file.
        docx.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create an ODT source document.
        // -----------------------------------------------------------------
        var odt = new Document();
        var builderOdt = new DocumentBuilder(odt);
        builderOdt.Writeln("This is content from the ODT source document.");
        // Save the ODT file.
        odt.Save(odtPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Load the source documents.
        // -----------------------------------------------------------------
        var srcDocx = new Document(docxPath);
        var srcOdt = new Document(odtPath);

        // -----------------------------------------------------------------
        // Prepare the destination document (start with the DOCX content).
        // -----------------------------------------------------------------
        // Clone the DOCX document to avoid modifying the original instance.
        // Clone returns a Node, so cast it back to Document.
        var destination = (Document)srcDocx.Clone(true);

        // Append the ODT document, preserving its original formatting.
        destination.AppendDocument(srcOdt, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the combined document as a PDF.
        // -----------------------------------------------------------------
        destination.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validate that the PDF was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The combined PDF file was not created.");

        Console.WriteLine($"Combined PDF created at: {Path.GetFullPath(pdfPath)}");
    }
}
