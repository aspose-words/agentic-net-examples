using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ReportGenerator
{
    public static void Main()
    {
        // Define file names for the sample DOCX and the resulting PDF.
        const string docxPath = "SampleInput.docx";
        const string pdfPath = "Report.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX containing placeholder tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Dear <<Name>>,");
        builder.Writeln("Thank you for your interest on <<Date>>.");
        builder.Writeln("Best regards,");
        builder.Writeln("Acme Corp.");
        templateDoc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX we just created.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // Step 3: Replace placeholders with actual values.
        // -----------------------------------------------------------------
        reportDoc.Range.Replace("<<Name>>", "John Doe");
        reportDoc.Range.Replace("<<Date>>", DateTime.Now.ToString("yyyy-MM-dd"));

        // -----------------------------------------------------------------
        // Step 4: Export the populated document to PDF.
        // -----------------------------------------------------------------
        reportDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 5: Validate that the PDF was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException($"The PDF file was not created: {pdfPath}");
        }

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(pdfPath)}");
    }
}
