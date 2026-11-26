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

        // Add a static header.
        builder.Writeln("=== Customer Report ===");
        builder.Writeln();

        // Insert merge fields for first and last name.
        builder.InsertField("MERGEFIELD FirstName");
        builder.Writeln(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln();

        // (Optional) Save the template for reference.
        string templatePath = "Template.docx";
        doc.Save(templatePath);

        // Prepare simple data source.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute mail merge with cleanup of unused fields.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the merge replaced the fields.
        string mergedText = doc.GetText();
        if (string.IsNullOrEmpty(mergedText) ||
            !mergedText.Contains("John") ||
            !mergedText.Contains("Doe"))
        {
            throw new InvalidOperationException("Mail merge failed: merged values not found in the document.");
        }

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Merged document was not created.", outputPath);
        }
    }
}
