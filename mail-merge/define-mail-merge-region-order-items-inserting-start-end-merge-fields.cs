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

            // Insert a mail merge region for order items.
            // The start field marks the beginning of the region.
            builder.InsertField(" MERGEFIELD TableStart:OrderItems");
            // Inside the region we place the fields that will be filled from the data source.
            builder.Write("Item: ");
            builder.InsertField(" MERGEFIELD ItemName");
            builder.Write(", Quantity: ");
            builder.InsertField(" MERGEFIELD Quantity");
            // The end field marks the end of the region.
            builder.InsertField(" MERGEFIELD TableEnd:OrderItems");

            // Build a DataTable that represents the OrderItems table.
            DataTable orderItems = new DataTable("OrderItems");
            orderItems.Columns.Add("ItemName", typeof(string));
            orderItems.Columns.Add("Quantity", typeof(int));

            // Add sample rows.
            orderItems.Rows.Add("Hawaiian Pizza", 2);
            orderItems.Rows.Add("Pepperoni Pizza", 1);
            orderItems.Rows.Add("Veggie Pizza", 3);

            // Execute the mail merge using the region.
            doc.MailMerge.ExecuteWithRegions(orderItems);

            // Save the resulting document.
            string outputPath = "OrderItemsMailMerge.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
