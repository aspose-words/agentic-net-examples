using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Build the template content with merge fields.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln(",");
        builder.Writeln("Your order number is ");
        builder.InsertField("MERGEFIELD OrderNumber");
        builder.Writeln(".");

        // Define a collection of records to merge.
        var records = new[]
        {
            new { FirstName = "John", LastName = "Doe", OrderNumber = "12345" },
            new { FirstName = "Jane", LastName = "Smith", OrderNumber = "67890" }
        };

        // Field names must match the merge fields in the template.
        string[] fieldNames = { "FirstName", "LastName", "OrderNumber" };

        // Ensure the output folder exists.
        string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // Process each record: clone the template, execute mail merge, and save.
        foreach (var record in records)
        {
            // Clone the template to keep it unchanged for the next iteration.
            Document doc = (Document)template.Clone(true);

            // Prepare the values for the current record.
            object[] fieldValues = { record.FirstName, record.LastName, record.OrderNumber };

            // Remove any unused fields after the merge.
            doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

            // Execute the simple mail merge.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Validate that the merge produced content and no unresolved fields remain.
            string mergedText = doc.GetText();
            if (string.IsNullOrWhiteSpace(mergedText) || mergedText.Contains("MERGEFIELD"))
                throw new InvalidOperationException("Mail merge failed for a record.");

            // Save the merged document to a separate file.
            string outputPath = Path.Combine(outputFolder, $"{record.FirstName}_{record.LastName}.docx");
            doc.Save(outputPath);
        }
    }
}
