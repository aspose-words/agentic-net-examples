using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert MERGEFIELD for the customer's name.
            builder.InsertField("MERGEFIELD CustomerName", "<CustomerName>");
            builder.Writeln(); // Move to the next line.

            // Insert MERGEFIELD for the customer's address.
            builder.InsertField("MERGEFIELD Address", "<Address>");
            builder.Writeln();

            // Prepare a data source with two columns: CustomerName and Address.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
            table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Execute the mail merge using the data table.
            doc.MailMerge.Execute(table);

            // Save the merged document to the file system.
            string outputPath = "MergedDocument.docx";
            doc.Save(outputPath);
        }
    }
}
