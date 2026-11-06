using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a page break.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");

        // Save the sample as PDF – this will be the input file.
        const string inputPdfPath = "input.pdf";
        sourceDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Load the PDF with options that ignore errors.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // By default, Aspose.Words ignores most loading errors.
            // If specific error‑ignoring flags existed, they would be set here.
        };
        Document pdfDoc = new Document(inputPdfPath, loadOptions);

        // Save the loaded PDF as DOCX while preserving page breaks.
        const string outputDocxPath = "output.docx";
        pdfDoc.Save(outputDocxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(outputDocxPath))
            throw new InvalidOperationException("The DOCX output file was not created.");

        // Optional cleanup (not required, but keeps the example tidy).
        // File.Delete(inputPdfPath);
        // File.Delete(outputDocxPath);
    }
}
