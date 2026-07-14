using System;
using System.Data;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple mail‑merge template with two fields: FirstName and LastName.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln();

        // Prepare a data source containing several records.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Rows.Add("John", "Doe");
        data.Rows.Add("Jane", "Smith");
        data.Rows.Add("Bob", "Johnson");

        // For each record clone the template, execute the mail merge, and save the result.
        int index = 1;
        foreach (DataRow row in data.Rows)
        {
            // Clone the original template so that each merge starts from a clean document.
            Document merged = (Document)template.Clone(true);

            // Perform mail merge for the current row.
            merged.MailMerge.Execute(row);

            // Save the merged document. Files will be named MergedDocument_1.docx, etc.
            string fileName = $"MergedDocument_{index}.docx";
            merged.Save(fileName);
            index++;
        }
    }
}
