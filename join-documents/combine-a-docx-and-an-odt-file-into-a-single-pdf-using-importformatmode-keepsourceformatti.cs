using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string odtPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.odt");
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.pdf");

        // -----------------------------------------------------------------
        // Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document docxDocument = new Document();
        DocumentBuilder docxBuilder = new DocumentBuilder(docxDocument);
        docxBuilder.Writeln("This is the content of the DOCX document.");
        docxDocument.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample ODT document.
        // -----------------------------------------------------------------
        Document odtDocument = new Document();
        DocumentBuilder odtBuilder = new DocumentBuilder(odtDocument);
        odtBuilder.Writeln("This is the content of the ODT document.");
        odtDocument.Save(odtPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Load the created documents.
        // -----------------------------------------------------------------
        Document srcDocx = new Document(docxPath);
        Document srcOdt = new Document(odtPath);

        // -----------------------------------------------------------------
        // Append the ODT document to the DOCX document, preserving its formatting.
        // -----------------------------------------------------------------
        srcDocx.AppendDocument(srcOdt, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the combined document as PDF.
        // -----------------------------------------------------------------
        srcDocx.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validate that the PDF file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Optional: output the location of the generated PDF.
        Console.WriteLine($"Merged PDF created at: {pdfPath}");
    }
}
