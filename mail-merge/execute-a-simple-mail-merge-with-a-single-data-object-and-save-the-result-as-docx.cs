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

        // Build a simple template with merge fields.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");
        builder.Writeln("Thank you for your interest.");

        // Prepare data for the mail merge.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute the mail merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = "MergedResult.docx";
        doc.Save(outputPath);

        // Validation: ensure the file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The merged document was not created.", outputPath);

        // Load the saved document to verify merged content.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        // Check that the merge fields were replaced.
        if (!resultText.Contains("John") || !resultText.Contains("Doe"))
            throw new InvalidOperationException("Mail merge failed: merged values not found in the output.");

        // Optionally, indicate success (no interactive output required).
        Console.WriteLine("Mail merge completed successfully. Output saved to " + outputPath);
    }
}
