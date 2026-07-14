using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        // Create a sample DOCX containing a date placeholder.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Report generated on {{Date}}.");
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document doc = new Document(inputPath);

        // Replace all occurrences of the placeholder with the current date.
        string placeholder = "{{Date}}";
        string currentDate = DateTime.Now.ToString("d");
        doc.Range.Replace(placeholder, currentDate, new FindReplaceOptions());

        // Export the updated document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF output file was not created.");
        }
    }
}
