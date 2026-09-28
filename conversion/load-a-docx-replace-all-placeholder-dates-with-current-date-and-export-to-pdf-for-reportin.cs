using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with a date placeholder.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Report generated on {{Date}}.");
        string inputPath = "input.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX.
        Document doc = new Document(inputPath);

        // Replace all occurrences of the placeholder with the current date.
        string placeholder = "{{Date}}";
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
        doc.Range.Replace(placeholder, currentDate, new FindReplaceOptions());

        // Export the document to PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
