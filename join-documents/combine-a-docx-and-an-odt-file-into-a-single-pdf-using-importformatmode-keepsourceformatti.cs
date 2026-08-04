using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string docxPath = Path.Combine(outputDir, "Sample.docx");
        string odtPath = Path.Combine(outputDir, "Sample.odt");
        string pdfPath = Path.Combine(outputDir, "Combined.pdf");

        // Create a sample DOCX document.
        Document docx = new Document();
        DocumentBuilder docxBuilder = new DocumentBuilder(docx);
        docxBuilder.Writeln("This is the DOCX part.");
        docx.Save(docxPath, SaveFormat.Docx);

        // Create a sample ODT document.
        Document odt = new Document();
        DocumentBuilder odtBuilder = new DocumentBuilder(odt);
        odtBuilder.Writeln("This is the ODT part.");
        odt.Save(odtPath, SaveFormat.Odt);

        // Load the created documents.
        Document srcDocx = new Document(docxPath);
        Document srcOdt = new Document(odtPath);

        // Append the ODT document to the DOCX document, preserving ODT formatting.
        srcDocx.AppendDocument(srcOdt, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF.
        srcDocx.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The combined PDF was not created.");

        // Optional: verify that both source texts are present in the PDF.
        Document pdfDoc = new Document(pdfPath);
        string pdfText = pdfDoc.GetText();

        if (!pdfText.Contains("DOCX part") || !pdfText.Contains("ODT part"))
            throw new InvalidOperationException("The combined PDF does not contain expected content.");
    }
}
