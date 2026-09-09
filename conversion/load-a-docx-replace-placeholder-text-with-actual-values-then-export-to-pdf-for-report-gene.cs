using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a placeholder.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear <<Name>>,");
        builder.Writeln("Thank you for your business.");
        const string templatePath = "template.docx";
        template.Save(templatePath, SaveFormat.Docx);

        // Step 2: Load the DOCX we just created.
        Document doc = new Document(templatePath);

        // Step 3: Replace the placeholder with an actual value.
        doc.Range.Replace("<<Name>>", "John Doe", new FindReplaceOptions());

        // Step 4: Export the populated document to PDF.
        const string pdfPath = "report.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Step 5: Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Optional: Clean up temporary files (not required, but keeps the folder tidy).
        File.Delete(templatePath);
    }
}
