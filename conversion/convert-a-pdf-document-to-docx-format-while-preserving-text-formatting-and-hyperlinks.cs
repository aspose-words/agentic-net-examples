using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pdfPath = "sample.pdf";
        const string docxPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample PDF with formatted text and a hyperlink.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Apply some text formatting.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Writeln("This is a sample paragraph with custom formatting.");

        // Insert a visible hyperlink.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", true);
        builder.Writeln(); // Add a line break after the hyperlink.

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // Step 2: Load the generated PDF and convert it to DOCX format.
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        pdfDoc.Save(docxPath, SaveFormat.Docx);

        // ---------------------------------------------------------------
        // Step 3: Validate that the DOCX file was created successfully.
        // ---------------------------------------------------------------
        if (!File.Exists(docxPath))
        {
            throw new InvalidOperationException($"The expected output file '{docxPath}' was not created.");
        }

        // Optional: Inform the user that conversion succeeded.
        Console.WriteLine($"PDF file '{pdfPath}' was successfully converted to DOCX file '{docxPath}'.");
    }
}
