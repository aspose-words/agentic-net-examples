using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with a placeholder for the date.
        Document sample = new Document();
        DocumentBuilder builder = new DocumentBuilder(sample);
        builder.Writeln("Report generated on {{Date}}.");
        const string inputPath = "input.docx";
        sample.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document doc = new Document(inputPath);

        // Replace all occurrences of the placeholder with the current date.
        string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
        doc.Range.Replace("{{Date}}", currentDate, new FindReplaceOptions(FindReplaceDirection.Forward));

        // Export the updated document to PDF.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");
    }
}
