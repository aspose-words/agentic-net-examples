using System;
using System.Data;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a mail‑merge template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert three merge fields: FirstName, LastName and Message.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");
        builder.Writeln();

        // Build a data source with several records.
        DataTable data = new DataTable("Recipients");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");

        data.Rows.Add("John", "Doe", "Hello! This is the first message.");
        data.Rows.Add("Jane", "Smith", "Greetings from the second record.");
        data.Rows.Add("Bob", "Johnson", "Your third message is here.");

        // For each record create a separate PDF file.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so each document starts from the same base.
            Document doc = (Document)template.Clone(true);

            // Perform mail merge for the current row only.
            doc.MailMerge.Execute(data.Rows[i]);

            // Save the merged document as a PDF file.
            string fileName = $"MergedDocument_{i + 1}.pdf";
            doc.Save(fileName, SaveFormat.Pdf);
        }
    }
}
