using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeRegionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a mail merge region named "OrderItems".
            // The region is marked by a TableStart field and a TableEnd field.
            builder.InsertField(" MERGEFIELD TableStart:OrderItems");
            builder.Writeln(); // Ensure the start field is on its own line.

            // Inside the region we place the fields that will be filled from the data source.
            builder.Write("Item: ");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.Write("\tQuantity: ");
            builder.InsertField(" MERGEFIELD Quantity");
            builder.Writeln(); // End of a row.

            // Insert the end field for the region.
            builder.InsertField(" MERGEFIELD TableEnd:OrderItems");
            builder.Writeln(); // Optional blank line after the region.

            // Build a DataTable that matches the region name and contains the data.
            DataTable orderItems = new DataTable("OrderItems");
            orderItems.Columns.Add("ItemName", typeof(string));
            orderItems.Columns.Add("Quantity", typeof(int));

            // Add some sample rows.
            orderItems.Rows.Add("Rugby Ball", 2);
            orderItems.Rows.Add("Team Jersey", 5);
            orderItems.Rows.Add("Water Bottle", 3);

            // Execute the mail merge with regions using the DataTable.
            doc.MailMerge.ExecuteWithRegions(orderItems);

            // Save the resulting document to disk.
            doc.Save("OrderItemsMailMerge.docx");
        }
    }
}
