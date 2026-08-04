using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX file containing placeholders.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Report for {{CustomerName}}");
        builder.Writeln("Date: {{ReportDate}}");
        const string inputFile = "input.docx";
        sourceDoc.Save(inputFile, SaveFormat.Docx);

        // Step 2: Load the DOCX file.
        Document doc = new Document(inputFile);

        // Step 3: Replace placeholders with actual values.
        doc.Range.Replace("{{CustomerName}}", "Acme Corp", new FindReplaceOptions(FindReplaceDirection.Forward));
        doc.Range.Replace("{{ReportDate}}", DateTime.Now.ToString("yyyy-MM-dd"), new FindReplaceOptions(FindReplaceDirection.Forward));

        // Step 4: Export the modified document to PDF.
        const string outputFile = "output.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Step 5: Verify that the PDF was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
