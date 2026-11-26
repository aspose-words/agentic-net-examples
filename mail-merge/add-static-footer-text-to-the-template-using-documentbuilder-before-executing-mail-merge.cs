using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the body with a merge field.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln(",");

        // Add static footer text before executing mail merge.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Confidential - For internal use only");

        // Return to the main story (optional, ensures further edits go to the body).
        builder.MoveToDocumentEnd();

        // Prepare simple mail merge data.
        string[] fieldNames = { "Name" };
        object[] fieldValues = { "John Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Clean up any unused fields.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output file was not created.");
        }

        // Load the saved document to verify that the merge succeeded.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("John Doe"))
        {
            throw new InvalidOperationException("Mail merge failed: merged value not found in the document.");
        }

        // Indicate successful completion.
        Console.WriteLine("Mail merge completed successfully.");
    }
}
