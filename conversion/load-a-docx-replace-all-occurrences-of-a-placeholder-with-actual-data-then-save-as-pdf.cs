using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file containing a placeholder.
        string inputPath = "input.docx";
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Hello <<Name>>! This is a sample document.");
        sourceDocument.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document document = new Document(inputPath);

        // Replace all occurrences of the placeholder with actual data.
        document.Range.Replace("<<Name>>", "John Doe", new FindReplaceOptions());

        // Save the modified document as PDF.
        string outputPath = "output.pdf";
        document.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
