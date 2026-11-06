using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string templatePath = "Template.docx";
        string mergedPath = "Merged.docx";
        string pdfConvertedPath = "PdfConverted.docx";
        string resultPath = "Result.docx";

        // 1. Create a DOCX template with merge fields.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report for: ");
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT");
        builder.Writeln();
        builder.Writeln("Date: ");
        builder.InsertField("MERGEFIELD ReportDate \\* MERGEFORMAT");
        builder.Writeln();
        builder.Writeln("This is the body of the mail‑merged document.");
        templateDoc.Save(templatePath, SaveFormat.Docx);

        // 2. Perform mail merge to generate a DOCX.
        var mergedDoc = new Document(templatePath);
        string[] fieldNames = { "Name", "ReportDate" };
        object[] fieldValues = { "John Doe", DateTime.Today.ToString("d") };
        mergedDoc.MailMerge.Execute(fieldNames, fieldValues);
        mergedDoc.Save(mergedPath, SaveFormat.Docx);

        // 3. Create a DOCX that simulates a PDF‑converted document.
        var pdfConvertedDoc = new Document();
        var pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        // Create a custom style in the destination document.
        Style destStyle = pdfConvertedDoc.Styles.Add(StyleType.Paragraph, "DestCustomStyle");
        destStyle.Font.Name = "Arial";
        destStyle.Font.Size = 14;
        destStyle.Font.Color = System.Drawing.Color.Blue;

        // Apply the custom style to some content.
        pdfBuilder.ParagraphFormat.Style = destStyle;
        pdfBuilder.Writeln("Original PDF‑converted content with custom style.");
        pdfBuilder.Writeln("Another paragraph in the destination document.");
        pdfConvertedDoc.Save(pdfConvertedPath, SaveFormat.Docx);

        // 4. Load both documents for joining.
        var destinationDoc = new Document(pdfConvertedPath);
        var sourceDoc = new Document(mergedPath);

        // 5. Append the mail‑merged document to the destination,
        // preserving destination styles.
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.UseDestinationStyles);

        // 6. Save the combined document.
        destinationDoc.Save(resultPath, SaveFormat.Docx);

        // 7. Validation: ensure the result file exists.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException($"Result file '{resultPath}' was not created.");

        // 8. Validation: ensure content from both source documents is present.
        var resultDoc = new Document(resultPath);
        bool hasPdfContent = resultDoc.GetText().Contains("Original PDF‑converted content");
        bool hasMergedContent = resultDoc.GetText().Contains("Report for:");
        if (!hasPdfContent || !hasMergedContent)
            throw new InvalidOperationException("The merged document does not contain expected content from both source files.");

        // Optional cleanup of intermediate files (can be omitted if inspection is needed).
        File.Delete(templatePath);
        File.Delete(mergedPath);
        File.Delete(pdfConvertedPath);
    }
}
