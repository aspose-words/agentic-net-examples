using System;
using System.Data;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a mail‑merge template document with three fields.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // -----------------------------------------------------------------
        // 2. Build a data source (DataTable) containing several records.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("Recipients");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");

        data.Rows.Add("John", "Doe", "Hello! This is the first merged document.");
        data.Rows.Add("Jane", "Smith", "Greetings from the second document.");
        data.Rows.Add("Bob", "Johnson", "This is the third merged file.");

        // -----------------------------------------------------------------
        // 3. For each record clone the template, perform a mail merge,
        //    and save the result as a separate file.
        // -----------------------------------------------------------------
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the original template so that each merge starts from a clean state.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current DataRow.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Save the merged document.
            string fileName = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");
            mergedDoc.Save(fileName);
        }

        // The program finishes without waiting for user input.
    }
}
