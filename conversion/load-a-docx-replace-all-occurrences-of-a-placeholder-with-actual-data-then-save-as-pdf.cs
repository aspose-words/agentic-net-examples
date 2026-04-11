using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        const string inputDocx = "SampleInput.docx";
        const string outputPdf = "Result.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX containing a placeholder.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Your order number is {{ORDER_NUMBER}}.");
        builder.Writeln("Thank you for shopping with us.");
        // Save the sample DOCX so that it can be loaded later.
        sampleDoc.Save(inputDocx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX, replace the placeholder, and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocx);

        // Replace all occurrences of the placeholder with actual data.
        int replacements = doc.Range.Replace("{{ORDER_NUMBER}}", "12345");
        // (Optional) You could verify that a replacement was made.
        if (replacements == 0)
        {
            throw new InvalidOperationException("Placeholder not found in the document.");
        }

        // Save the modified document as PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdf))
        {
            throw new FileNotFoundException($"The PDF file '{outputPdf}' was not created.");
        }

        // The example finishes without requiring any user interaction.
    }
}
