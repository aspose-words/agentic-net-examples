using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a date placeholder.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Report generated on {{Date}}.");
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Load the DOCX.
        Document doc = new Document(inputPath);

        // Step 3: Replace all date placeholders with the current date.
        string placeholder = "{{Date}}";
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
        doc.Range.Replace(placeholder, currentDate, new FindReplaceOptions());

        // Step 4: Export the document to PDF.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Step 5: Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
