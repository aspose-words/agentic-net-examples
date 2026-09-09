using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        // -----------------------------------------------------------------
        // Create a sample DOCX file containing a placeholder.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear _Placeholder_,");
        builder.Writeln("Thank you for using Aspose.Words.");
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the DOCX, replace the placeholder, and save as PDF.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);
        const string placeholder = "_Placeholder_";
        const string actualData = "John Doe";

        // Replace all occurrences of the placeholder.
        doc.Range.Replace(placeholder, actualData);

        // Save the modified document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }
    }
}
