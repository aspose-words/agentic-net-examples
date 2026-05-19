using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;   // Needed for FindReplaceOptions

public class ReportGenerator
{
    public static void Main()
    {
        // Define file names.
        const string inputDocx = "template.docx";
        const string outputPdf = "report.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX with placeholder fields.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Dear {CustomerName},");
        builder.Writeln("Your order number {OrderNumber} has been shipped.");
        builder.Writeln("Thank you for shopping with us.");
        template.Save(inputDocx, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX, replace placeholders with actual values.
        // -----------------------------------------------------------------
        Document doc = new Document(inputDocx);
        doc.Range.Replace("{CustomerName}", "John Doe", new FindReplaceOptions());
        doc.Range.Replace("{OrderNumber}", "A12345", new FindReplaceOptions());

        // -----------------------------------------------------------------
        // Step 3: Export the populated document to PDF.
        // -----------------------------------------------------------------
        doc.Save(outputPdf, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 4: Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF report was not created.");

        // Optional: clean up temporary files (comment out if you need the files later).
        // File.Delete(inputDocx);
    }
}
