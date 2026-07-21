using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Define file names for the intermediate PDF and the final DOCX.
        const string pdfPath = "sample.pdf";
        const string docxPath = "converted.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document and save it as a PDF file.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF created by Aspose.Words.");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // ---------------------------------------------------------------
        // Step 2: Load the PDF using PdfLoadOptions (no password supplied).
        // ---------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions(); // No password set; password protection is ignored.
        Document pdfDoc = new Document(pdfPath, loadOptions);

        // ---------------------------------------------------------------
        // Step 3: Save the loaded document as DOCX.
        // ---------------------------------------------------------------
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // Validate that the DOCX file was created.
        if (!File.Exists(docxPath))
            throw new InvalidOperationException("The DOCX conversion output was not created.");

        // Optional: Inform the user that the conversion succeeded.
        Console.WriteLine($"PDF '{pdfPath}' was successfully converted to DOCX '{docxPath}'.");
    }
}
