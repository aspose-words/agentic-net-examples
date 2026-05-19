using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the intermediate DOCX and final PDF.
        const string docxPath = "sample.docx";
        const string pdfPath = "output.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX file containing a placeholder.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Thank you for using Aspose.Words.");
        // Save the document as DOCX so it can be loaded later.
        sampleDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX, replace the placeholder, and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(docxPath);
        // Replace all occurrences of the placeholder with actual data.
        int replacements = doc.Range.Replace("{{Name}}", "John Doe");
        if (replacements == 0)
            throw new InvalidOperationException("Placeholder was not found in the document.");

        // Save the modified document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 3: Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, clean up the intermediate DOCX file.
        // File.Delete(docxPath);
    }
}
