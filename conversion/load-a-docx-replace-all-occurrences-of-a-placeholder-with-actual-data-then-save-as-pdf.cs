using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file with a placeholder.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Dear {{NAME}},");
        builder.Writeln("Thank you for using our service.");
        sampleDoc.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the DOCX file.
        Document doc = new Document("input.docx");

        // Step 3: Replace all occurrences of the placeholder with actual data.
        string placeholder = "{{NAME}}";
        string actualData = "John Doe";
        doc.Range.Replace(placeholder, actualData, new FindReplaceOptions(FindReplaceDirection.Forward));

        // Step 4: Save the modified document as PDF.
        string outputPath = "output.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Step 5: Validate that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
