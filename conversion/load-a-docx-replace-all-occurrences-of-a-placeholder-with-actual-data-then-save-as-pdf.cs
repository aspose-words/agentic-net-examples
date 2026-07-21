using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file with a placeholder.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample document containing a placeholder: _PLACEHOLDER_.");
        source.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the DOCX file.
        Document doc = new Document("input.docx");

        // Step 3: Replace the placeholder with actual data.
        int replacements = doc.Range.Replace("_PLACEHOLDER_", "Actual Data");
        Console.WriteLine($"Number of replacements made: {replacements}");

        // Step 4: Save the modified document as PDF.
        doc.Save("output.pdf", SaveFormat.Pdf);

        // Step 5: Verify that the PDF was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
