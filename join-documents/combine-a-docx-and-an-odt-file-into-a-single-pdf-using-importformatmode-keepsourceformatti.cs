using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample files and the final PDF.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string docxPath = Path.Combine(outputDir, "SourceDocument.docx");
        string odtPath = Path.Combine(outputDir, "SourceDocument.odt");
        string pdfPath = Path.Combine(outputDir, "MergedDocument.pdf");

        // -----------------------------------------------------------------
        // Create a DOCX source document.
        // -----------------------------------------------------------------
        Document docx = new Document();
        DocumentBuilder docxBuilder = new DocumentBuilder(docx);
        docxBuilder.Writeln("This is content from the DOCX source document.");
        docx.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create an ODT source document.
        // -----------------------------------------------------------------
        Document odt = new Document();
        DocumentBuilder odtBuilder = new DocumentBuilder(odt);
        odtBuilder.Writeln("This is content from the ODT source document.");
        // Save as ODT using OdtSaveOptions.
        OdtSaveOptions odtSaveOptions = new OdtSaveOptions();
        odt.Save(odtPath, odtSaveOptions);

        // -----------------------------------------------------------------
        // Load the two documents.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(docxPath);
        Document odtDoc = new Document(odtPath);

        // Append the ODT document to the DOCX document, preserving its formatting.
        mainDoc.AppendDocument(odtDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF.
        mainDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("Documents merged and saved to PDF successfully.");
    }
}
