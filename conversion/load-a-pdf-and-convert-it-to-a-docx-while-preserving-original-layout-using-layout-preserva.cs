using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample document with some content.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF document.");
        builder.InsertParagraph();
        builder.Writeln("It contains multiple lines of text to demonstrate layout preservation.");
        builder.InsertParagraph();
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndTable();

        // Save the sample document as PDF.
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF. Layout preservation is enabled by default for Aspose.Words PDF loading.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Convert the loaded PDF to DOCX.
        string docxPath = "output.docx";
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException("The DOCX file was not created as expected.");
        }

        // Optional clean‑up (commented out to allow inspection of the files after the run).
        // File.Delete(pdfPath);
        // File.Delete(docxPath);
    }
}
