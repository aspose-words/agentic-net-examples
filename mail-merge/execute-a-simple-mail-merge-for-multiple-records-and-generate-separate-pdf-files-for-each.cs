using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare the data source – a DataTable with several records.
        DataTable table = new DataTable("Recipients");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Message");

        table.Rows.Add(new object[] { "John", "Doe", "Hello John! This is your personalized message." });
        table.Rows.Add(new object[] { "Jane", "Smith", "Hi Jane, welcome to our service." });
        table.Rows.Add(new object[] { "Bob", "Johnson", "Dear Bob, thank you for your purchase." });

        // Build the mail‑merge template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MailMergeOutputs");
        Directory.CreateDirectory(outputDir);

        // For each record create a separate PDF file.
        int index = 1;
        foreach (DataRow row in table.Rows)
        {
            // Clone the template so each document starts from the same base.
            Document doc = (Document)template.Clone();

            // Perform mail merge for the current row only.
            doc.MailMerge.Execute(row);

            // Save the result as a PDF file.
            string outPath = Path.Combine(outputDir, $"MergedDocument_{index}.pdf");
            doc.Save(outPath, SaveFormat.Pdf);

            index++;
        }
    }
}
