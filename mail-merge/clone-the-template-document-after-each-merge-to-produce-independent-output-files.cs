using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeCloneExample
{
    public static void Main()
    {
        // Create a simple mail‑merge template in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare sample data – a DataTable with two records.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello! This is the first merged document.");
        data.Rows.Add("Jane", "Smith", "Greetings from the second merged document.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutputs");
        Directory.CreateDirectory(outputDir);

        // Perform a separate mail merge for each row, cloning the template each time.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the original template to keep it unchanged for the next iteration.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current DataRow.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Save the merged document to a distinct file.
            string outputPath = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");
            mergedDoc.Save(outputPath);
        }
    }
}
