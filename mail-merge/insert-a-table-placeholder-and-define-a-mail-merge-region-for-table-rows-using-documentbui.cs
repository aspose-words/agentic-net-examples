using System;
using System.Data;
using Aspose.Words;

namespace MailMergeTableRegionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a title for the table.
            builder.Writeln("Orders:");

            // Insert the start of a mail merge region named "Orders".
            builder.InsertField(" MERGEFIELD TableStart:Orders ");

            // Build a table that will be repeated for each record in the data source.
            builder.StartTable();

            // First column – Item name.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Item ");

            // Second column – Quantity.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity ");

            // End the row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Insert the end of the mail merge region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders ");

            // Prepare a DataTable that matches the region name and contains the data.
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("Item");
            ordersTable.Columns.Add("Quantity");

            ordersTable.Rows.Add("Apple", 5);
            ordersTable.Rows.Add("Banana", 12);
            ordersTable.Rows.Add("Orange", 8);

            // Execute the mail merge with regions – the table rows will be duplicated for each record.
            doc.MailMerge.ExecuteWithRegions(ordersTable);

            // Save the result to a file in the same folder as the executable.
            string outputPath = "MailMergeTableRegion.docx";
            doc.Save(outputPath);
        }
    }
}
