using System;
using System.Data;
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

        // Define the region name for the table.
        const string regionName = "Employees";

        // Insert the TableStart field.
        builder.InsertField($"MERGEFIELD TableStart:{regionName}", null);
        builder.Writeln();

        // Build a table with merge fields for each column.
        builder.StartTable();

        // Header row (optional, not a merge field).
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.EndRow();

        // Data row with merge fields.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Name", null);
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Age", null);
        builder.EndRow();

        builder.EndTable();

        // Insert the TableEnd field.
        builder.InsertField($"MERGEFIELD TableEnd:{regionName}", null);
        builder.Writeln();

        // Prepare the data source as a DataTable.
        DataTable table = new DataTable(regionName);
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Age", typeof(int));

        table.Rows.Add("John Doe", 30);
        table.Rows.Add("Jane Smith", 25);
        table.Rows.Add("Bob Johnson", 40);

        // Configure mail merge cleanup to remove unused fields.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Execute the region-based mail merge.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
        doc.Save(outputPath);

        // Validate that the merged output contains the expected data.
        string docText = doc.GetText();
        if (!docText.Contains("John Doe") || !docText.Contains("Jane Smith") || !docText.Contains("Bob Johnson"))
        {
            throw new InvalidOperationException("Merged document does not contain expected data.");
        }

        // Indicate successful completion.
        Console.WriteLine("Mail merge completed successfully. Output saved to: " + outputPath);
    }
}
