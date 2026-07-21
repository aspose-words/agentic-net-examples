using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX with placeholder tokens.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Report for {Name}");
        builder.Writeln("Date: {Date}");
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX, replace placeholders with actual values.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);
        doc.Range.Replace("{Name}", "John Doe", new FindReplaceOptions());
        doc.Range.Replace("{Date}", DateTime.Today.ToString("yyyy-MM-dd"), new FindReplaceOptions());

        // -----------------------------------------------------------------
        // Step 3: Export the modified document to PDF.
        // -----------------------------------------------------------------
        doc.Save(outputPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: Ensure the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
