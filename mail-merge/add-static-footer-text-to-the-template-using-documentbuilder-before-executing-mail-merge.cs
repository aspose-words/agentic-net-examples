using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the body of the template with a merge field.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln(",");

        // Add a static footer text before executing the mail merge.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Writeln("Company Confidential - Do not distribute");

        // Return to the main story (optional, not required for this simple example).
        builder.MoveToDocumentEnd();

        // Prepare simple data source for the mail merge.
        string[] fieldNames = { "Name" };
        object[] fieldValues = { "John Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Validation: ensure the merged field was replaced.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("John Doe"))
            throw new Exception("Mail merge failed – the field value was not inserted.");

        // Validation: ensure the static footer text is present.
        if (!mergedText.Contains("Company Confidential"))
            throw new Exception("Static footer text is missing.");

        // Indicate successful completion (optional, not required for non‑interactive run).
        Console.WriteLine("Mail merge completed successfully.");
    }
}
