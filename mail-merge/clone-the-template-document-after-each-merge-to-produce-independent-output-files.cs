using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a mail‑merge template in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare sample data.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello! This is a merged document.");
        data.Rows.Add("Jane", "Smith", "Welcome to Aspose.Words mail merge.");

        // Output folder (creates it if it does not exist).
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutputs");
        Directory.CreateDirectory(outputDir);

        // Perform a separate merge for each record, cloning the template each time.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the original template so each output file is independent.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current DataRow.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Save the result to a distinct file.
            string fileName = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");
            mergedDoc.Save(fileName);
        }
    }
}
