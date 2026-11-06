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
        builder.Writeln("=== Invoice ===");
        builder.Writeln();

        // Insert merge fields for Name and Date.
        builder.Write("Hello ");
        builder.InsertField("MERGEFIELD Name");
        builder.Write(", your appointment is on ");
        builder.InsertField("MERGEFIELD Date");
        builder.Writeln(".");

        // Ensure unused fields are removed after merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Prepare simple data source.
        string[] fieldNames = { "Name", "Date" };
        string dateString = DateTime.Now.ToString("D");
        object[] fieldValues = { "John Doe", dateString };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the merge replaced the fields.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("John Doe") || !mergedText.Contains(dateString))
        {
            throw new InvalidOperationException("Mail merge did not replace all fields correctly.");
        }

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The merged document was not created.", outputPath);
        }
    }
}
