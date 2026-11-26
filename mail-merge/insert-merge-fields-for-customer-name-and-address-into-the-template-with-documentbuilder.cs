using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the merged output.
        string templatePath = Path.Combine(Path.GetTempPath(), "Template.docx");
        string outputPath = Path.Combine(Path.GetTempPath(), "MergedDocument.docx");

        // Create a new document and insert merge fields using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Customer Information:");
        builder.InsertField("MERGEFIELD CustomerName");
        builder.Writeln();
        builder.InsertField("MERGEFIELD CustomerAddress");
        builder.Writeln();
        doc.Save(templatePath);

        // Load the template document.
        Document templateDoc = new Document(templatePath);

        // Define field names and corresponding values.
        string[] fieldNames = { "CustomerName", "CustomerAddress" };
        object[] fieldValues = { "John Doe", "123 Main St, Anytown" };

        // Perform the mail merge.
        templateDoc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        templateDoc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        templateDoc.Save(outputPath);

        // Validate that the merged output contains the expected values.
        string mergedText = templateDoc.GetText();
        if (!mergedText.Contains("John Doe") || !mergedText.Contains("123 Main St, Anytown"))
        {
            throw new InvalidOperationException("Mail merge failed: merged values not found in the document.");
        }

        // Ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Merged document was not created.", outputPath);
        }

        // Optional: clean up temporary files (comment out if you want to inspect the files).
        // File.Delete(templatePath);
        // File.Delete(outputPath);
    }
}
