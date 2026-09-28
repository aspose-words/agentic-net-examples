using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple merge field for the order ID.
        builder.Writeln("Order ID:");
        builder.InsertField("MERGEFIELD OrderID \\* MERGEFORMAT");
        builder.Writeln();

        // Insert the start of the mail merge region for order items.
        builder.InsertField("MERGEFIELD TableStart:OrderItems \\* MERGEFORMAT");
        builder.Writeln();

        // Create a table that will be repeated for each item in the region.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item Name");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data row with merge fields.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD ItemName \\* MERGEFORMAT");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity \\* MERGEFORMAT");
        builder.EndRow();

        builder.EndTable();

        // Insert the end of the mail merge region.
        builder.InsertField("MERGEFIELD TableEnd:OrderItems \\* MERGEFORMAT");
        builder.Writeln();

        // Prepare the data for the merge.
        var order = new Order
        {
            OrderID = "12345",
            Items = new List<Item>
            {
                new Item { ItemName = "Apple",  Quantity = 5 },
                new Item { ItemName = "Banana", Quantity = 3 },
                new Item { ItemName = "Orange", Quantity = 7 }
            }
        };

        // Perform a simple merge for the OrderID field.
        doc.MailMerge.Execute(new[] { "OrderID" }, new object[] { order.OrderID });

        // Convert the collection of items to a DataTable for region merge.
        DataTable itemsTable = new DataTable("OrderItems");
        itemsTable.Columns.Add("ItemName", typeof(string));
        itemsTable.Columns.Add("Quantity", typeof(int));

        foreach (var item in order.Items)
        {
            itemsTable.Rows.Add(item.ItemName, item.Quantity);
        }

        // Perform a region merge for the collection of order items.
        doc.MailMerge.ExecuteWithRegions(itemsTable);

        // Save the merged document.
        doc.Save("MergedOutput.docx");
    }

    // Simple class representing an order.
    public class Order
    {
        public string OrderID { get; set; }
        public List<Item> Items { get; set; }
    }

    // Simple class representing an individual order item.
    public class Item
    {
        public string ItemName { get; set; }
        public int Quantity { get; set; }
    }
}
