using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a placeholder.
        Document sourceDocument = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDocument);
        builder.Writeln("Report for {{Name}}");
        sourceDocument.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the DOCX and replace the placeholder with actual value.
        Document document = new Document("input.docx");
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        document.Range.Replace("{{Name}}", "John Doe", replaceOptions);

        // Step 3: Export the modified document to PDF.
        document.Save("output.pdf", SaveFormat.Pdf);

        // Step 4: Validate that the PDF was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }

        // Optional: Clean up created files (comment out if you want to keep them).
        // File.Delete("input.docx");
        // File.Delete("output.pdf");
    }
}
