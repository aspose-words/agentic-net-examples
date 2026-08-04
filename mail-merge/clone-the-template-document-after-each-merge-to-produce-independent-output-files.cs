using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class MailMergeCloneExample
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document in memory with three merge fields.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare a data source with several records.
        DataTable data = new DataTable("Recipients");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello! This is the first message.");
        data.Rows.Add("Jane", "Smith", "Greetings from the second record.");
        data.Rows.Add("Bob", "Johnson", "Third message goes here.");

        // Iterate over each row, clone the template, perform mail merge, and save the result.
        int index = 1;
        foreach (DataRow row in data.Rows)
        {
            // Clone the template to obtain an independent document for this record.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current row.
            mergedDoc.MailMerge.Execute(row);

            // Save the merged document to a uniquely named file.
            string fileName = Path.Combine(outputDir, $"MergedDocument_{index}.docx");
            mergedDoc.Save(fileName);
            index++;
        }
    }
}
