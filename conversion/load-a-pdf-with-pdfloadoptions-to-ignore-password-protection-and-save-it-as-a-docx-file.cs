using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and the final DOCX output.
        string pdfPath = "sample.pdf";
        string docxPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a simple PDF document.
        // -----------------------------------------------------------------
        Document creator = new Document();
        DocumentBuilder builder = new DocumentBuilder(creator);
        builder.Writeln("Hello from PDF created by Aspose.Words.");
        creator.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 2: Load the PDF using PdfLoadOptions.
        //         No password is supplied, effectively ignoring any password.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Step 3: Save the loaded document as DOCX.
        // -----------------------------------------------------------------
        pdfDocument.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Validation: Ensure the DOCX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException($"Failed to create output file '{docxPath}'.");
        }

        // Indicate successful completion.
        Console.WriteLine("Conversion completed successfully.");
    }
}
