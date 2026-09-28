using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string templatePath = "MailMergeTemplate.docx";
        string mailMergeResultPath = "MailMergeResult.docx";
        string pdfConvertedDocPath = "PdfConverted.docx";
        string mergedOutputPath = "MergedDocument.docx";

        // --------------------------------------------------------------------
        // 1. Create a DOCX template with proper MERGEFIELD fields.
        // --------------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Write("Report for: ");
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        builder.Writeln();
        builder.Write("Date: ");
        builder.InsertField("MERGEFIELD Date \\* MERGEFORMAT");
        builder.Writeln();
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // --------------------------------------------------------------------
        // 2. Perform mail merge to generate a DOCX.
        // --------------------------------------------------------------------
        var mailMergeDoc = new Document(templatePath);
        var mergeData = new
        {
            Name = "John Doe",
            Date = DateTime.Today.ToString("d")
        };
        mailMergeDoc.MailMerge.Execute(
            new[] { "Name", "Date" },
            new object[] { mergeData.Name, mergeData.Date });
        mailMergeDoc.Save(mailMergeResultPath, SaveFormat.Docx);

        // --------------------------------------------------------------------
        // 3. Create a DOCX that simulates a PDF‑converted document.
        // --------------------------------------------------------------------
        var pdfConvertedDoc = new Document();
        var pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        pdfBuilder.Font.Name = "Arial";
        pdfBuilder.Font.Size = 14;
        pdfBuilder.Writeln("PDF Converted Content");
        pdfBuilder.Writeln("This text represents a document originally created from a PDF.");
        pdfConvertedDoc.Save(pdfConvertedDocPath, SaveFormat.Docx);

        // --------------------------------------------------------------------
        // 4. Load the destination (PDF‑converted) document.
        // --------------------------------------------------------------------
        var destinationDoc = new Document(pdfConvertedDocPath);

        // --------------------------------------------------------------------
        // 5. Load the mail‑merged source document.
        // --------------------------------------------------------------------
        var sourceDoc = new Document(mailMergeResultPath);

        // --------------------------------------------------------------------
        // 6. Append the source document to the destination,
        //    preserving the destination's styles.
        // --------------------------------------------------------------------
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

        // --------------------------------------------------------------------
        // 7. Save the merged document.
        // --------------------------------------------------------------------
        destinationDoc.Save(mergedOutputPath, SaveFormat.Docx);

        // --------------------------------------------------------------------
        // 8. Validation: ensure the file exists and contains content from both sources.
        // --------------------------------------------------------------------
        if (!File.Exists(mergedOutputPath))
        {
            throw new InvalidOperationException($"Merged file was not created at '{mergedOutputPath}'.");
        }

        var mergedDoc = new Document(mergedOutputPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("PDF Converted Content"))
        {
            throw new InvalidOperationException("Merged document is missing content from the PDF‑converted source.");
        }

        if (!mergedText.Contains("John Doe"))
        {
            throw new InvalidOperationException("Merged document is missing content from the mail‑merged source.");
        }

        // Program completed successfully.
    }
}
