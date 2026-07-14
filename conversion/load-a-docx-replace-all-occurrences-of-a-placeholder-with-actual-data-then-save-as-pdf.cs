using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a placeholder.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document containing a placeholder: _PLACEHOLDER_.");

        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Load the DOCX, replace the placeholder, and save as PDF.
        Document doc = new Document(inputPath);
        int replacements = doc.Range.Replace("_PLACEHOLDER_", "Actual Data");

        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Step 3: Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output file was not created.");

        // Optional: output the number of replacements performed.
        Console.WriteLine($"Replacements made: {replacements}");
    }
}
