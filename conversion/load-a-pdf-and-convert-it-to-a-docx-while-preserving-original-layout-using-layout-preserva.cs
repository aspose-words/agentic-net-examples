using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.pdf");
        string docxPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleConverted.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple PDF document using Aspose.Words.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfSource);
        builder.Writeln("This is a sample PDF created with Aspose.Words.");
        builder.Writeln("The layout of this document will be preserved when converting to DOCX.");
        // Save the document as PDF.
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new FileNotFoundException("Failed to create the source PDF file.", pdfPath);

        // -----------------------------------------------------------------
        // 2. Load the PDF. No special load options are required for layout preservation.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions(); // default options
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Convert the loaded PDF to DOCX while preserving layout.
        // -----------------------------------------------------------------
        // OoxmlSaveOptions can be used for DOCX output; no extra settings are needed.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        pdfDocument.Save(docxPath, saveOptions);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("PDF to DOCX conversion failed; output file not found.", docxPath);

        // Indicate successful completion.
        Console.WriteLine("PDF successfully converted to DOCX with layout preservation.");
    }
}
