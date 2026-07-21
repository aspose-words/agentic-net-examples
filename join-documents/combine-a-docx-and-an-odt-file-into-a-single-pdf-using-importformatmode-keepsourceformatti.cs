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
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Combined.pdf");

        // -----------------------------------------------------------------
        // Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document docx = new Document();
        DocumentBuilder builder = new DocumentBuilder(docx);
        builder.Writeln("This is the content of the DOCX document.");
        docx.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample ODT document.
        // -----------------------------------------------------------------
        Document odt = new Document();
        builder = new DocumentBuilder(odt);
        builder.Writeln("This is the content of the ODT document.");
        odt.Save(odtPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Load the source documents.
        // -----------------------------------------------------------------
        Document srcDocx = new Document(docxPath);
        Document srcOdt = new Document(odtPath);

        // -----------------------------------------------------------------
        // Create the destination document and append the sources.
        // The ODT content is appended with KeepSourceFormatting.
        // -----------------------------------------------------------------
        Document dst = new Document();

        // Append DOCX (default formatting behavior).
        dst.AppendDocument(srcDocx, ImportFormatMode.UseDestinationStyles);

        // Append ODT while preserving its original formatting.
        dst.AppendDocument(srcOdt, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the combined document as PDF.
        // -----------------------------------------------------------------
        dst.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        // Cleanup temporary files (optional).
        File.Delete(docxPath);
        File.Delete(odtPath);
    }
}
