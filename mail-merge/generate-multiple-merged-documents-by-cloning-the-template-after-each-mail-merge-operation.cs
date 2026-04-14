using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a mail‑merge template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert three merge fields that will be filled from the data source.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare a DataTable that holds the data for several merged documents.
        DataTable data = new DataTable("Recipients");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");

        data.Rows.Add("John", "Doe", "Hello! This is the first merged document.");
        data.Rows.Add("Jane", "Smith", "Greetings from the second merged document.");
        data.Rows.Add("Bob", "Johnson", "This is the third merged document.");

        // Directory where the merged documents will be saved.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocs");
        Directory.CreateDirectory(outputDir);

        // Iterate over each data row, clone the template, perform mail merge, and save.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template to obtain a fresh document for this record.
            Document mergedDoc = (Document)template.Clone(true);

            // Execute mail merge for the current DataRow.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Build a file name such as "MergedDocument_1.docx".
            string fileName = Path.Combine(outputDir, $"MergedDocument_{i + 1}.docx");

            // Save the merged document.
            mergedDoc.Save(fileName);
        }

        // The program finishes automatically; no user interaction is required.
    }
}
