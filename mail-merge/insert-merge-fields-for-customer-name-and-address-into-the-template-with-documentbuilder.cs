using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a greeting.
            builder.Write("Dear ");

            // Insert a MERGEFIELD for the customer's name.
            builder.InsertField("MERGEFIELD CustomerName", "<CustomerName>");

            builder.Write(",\n");
            builder.Writeln("Your address is:");

            // Insert a MERGEFIELD for the customer's address.
            builder.InsertField("MERGEFIELD Address", "<Address>");

            // Prepare sample data in a DataTable.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
            table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Execute the mail merge using the DataTable as the data source.
            doc.MailMerge.Execute(table);

            // Save the resulting document.
            doc.Save("MergedCustomers.docx");
        }
    }
}
