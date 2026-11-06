using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the template with merge fields for customer name and address.
        builder.Writeln("Customer Information:");
        builder.InsertField("MERGEFIELD CustomerName");
        builder.Writeln();
        builder.InsertField("MERGEFIELD CustomerAddress");
        builder.Writeln();

        // Prepare simple data source.
        string[] fieldNames = new string[] { "CustomerName", "CustomerAddress" };
        object[] fieldValues = new object[] { "John Doe", "123 Main St, Anytown" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Remove any unused merge fields.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Save the merged document.
        string outputPath = "MergedDocument.docx";
        doc.Save(outputPath);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged document was not created.");

        // Load the saved document to verify merged content.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        // Ensure that the merge fields were replaced with the provided values.
        if (!resultText.Contains("John Doe") || !resultText.Contains("123 Main St"))
            throw new InvalidOperationException("Mail merge failed: expected values not found in the document.");

        // Optionally, clean up the output file after verification (comment out if you want to keep it).
        // File.Delete(outputPath);
    }
}
