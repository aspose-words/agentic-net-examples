using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a placeholder.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Report for {{Name}}");
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath, SaveFormat.Docx);

        // Step 2: Load the DOCX.
        Document doc = new Document(inputPath);

        // Step 3: Replace the placeholder with actual value.
        const string placeholder = "{{Name}}";
        const string actualValue = "John Doe";
        doc.Range.Replace(placeholder, actualValue, new FindReplaceOptions(FindReplaceDirection.Forward));

        // Step 4: Export to PDF.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Step 5: Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: Clean up sample input file (not required, but keeps the folder tidy).
        // File.Delete(inputPath);
    }
}
