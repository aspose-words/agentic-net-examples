using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutputs");
        Directory.CreateDirectory(outputDir);

        // Create a simple mail‑merge template in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Build a data source with several records.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Welcome to our service!");
        data.Rows.Add("Jane", "Smith", "Your order has shipped.");
        data.Rows.Add("Bob", "Johnson", "Thank you for your feedback.");

        // Perform a separate merge for each record, cloning the template each time.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Deep clone the template so each output file is independent.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current DataRow.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Save the merged document with a unique name.
            string outFile = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");
            mergedDoc.Save(outFile);
        }
    }
}
