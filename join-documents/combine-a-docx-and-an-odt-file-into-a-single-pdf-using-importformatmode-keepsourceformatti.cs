using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Define file paths.
        string docxPath = Path.Combine(outputDir, "sample.docx");
        string odtPath = Path.Combine(outputDir, "sample.odt");
        string pdfPath = Path.Combine(outputDir, "merged.pdf");

        // Create a DOCX document with sample content.
        Document docx = new Document();
        docx.FirstSection.Body.FirstParagraph.AppendChild(new Run(docx, "This is DOCX content."));
        docx.Save(docxPath, SaveFormat.Docx);

        // Create an ODT document with sample content.
        Document odt = new Document();
        odt.FirstSection.Body.FirstParagraph.AppendChild(new Run(odt, "This is ODT content."));
        odt.Save(odtPath, SaveFormat.Odt);

        // Load the created documents.
        Document destination = new Document(docxPath);
        Document sourceOdt = new Document(odtPath);

        // Append the ODT document to the DOCX document, preserving its formatting.
        destination.AppendDocument(sourceOdt, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as a PDF.
        destination.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Indicate successful completion.
        Console.WriteLine($"Merged PDF created at: {pdfPath}");
    }
}
