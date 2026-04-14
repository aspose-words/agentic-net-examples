using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeRegionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a mail merge region named "Orders".
            // The region is marked by TableStart and TableEnd merge fields.
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Writeln(); // Ensure the fields are on separate lines.

            // Insert fields that will be populated for each order item.
            builder.Write("Item: ");
            builder.InsertField(" MERGEFIELD Item");
            builder.Writeln();

            builder.Write("Quantity: ");
            builder.InsertField(" MERGEFIELD Quantity");
            builder.Writeln();

            // End of the "Orders" region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.Writeln();

            // Prepare a DataTable that matches the region name and contains order data.
            DataTable ordersTable = new DataTable("Orders");
            ordersTable.Columns.Add("Item");
            ordersTable.Columns.Add("Quantity");

            ordersTable.Rows.Add(new object[] { "Rugby Ball", 2 });
            ordersTable.Rows.Add(new object[] { "Team Jersey", 5 });
            ordersTable.Rows.Add(new object[] { "Water Bottle", 3 });

            // Execute the mail merge with regions using the DataTable.
            doc.MailMerge.ExecuteWithRegions(ordersTable);

            // Save the resulting document.
            doc.Save("MailMergeRegion.docx");
        }
    }
}
