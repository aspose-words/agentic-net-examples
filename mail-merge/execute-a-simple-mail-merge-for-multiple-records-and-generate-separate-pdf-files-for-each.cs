using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a mail‑merge template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert merge fields that will be filled with data.
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

        data.Rows.Add("John", "Doe", "Hello! This is your first message.");
        data.Rows.Add("Jane", "Smith", "Welcome to the Aspose.Words demo.");
        data.Rows.Add("Bob", "Johnson", "Your order has been shipped.");

        // For each record create a separate document, perform the mail merge and save as PDF.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so each document starts from the same base.
            Document doc = (Document)template.Clone(true);

            // Merge the current row into the document.
            doc.MailMerge.Execute(data.Rows[i]);

            // Save the result as a PDF file.
            string fileName = $"MergedDocument_{i + 1}.pdf";
            doc.Save(fileName, SaveFormat.Pdf);
        }
    }
}
