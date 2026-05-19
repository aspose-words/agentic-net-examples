using System;
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

        // Add a heading.
        builder.Writeln("Order Summary");
        builder.Writeln();

        // Insert the start of a mail merge region named "OrderItems".
        // The region is defined by a MERGEFIELD with the prefix TableStart:.
        builder.InsertField(" MERGEFIELD TableStart:OrderItems");

        // Fields that will be populated for each row of the data source.
        builder.Write("Item: ");
        builder.InsertField(" MERGEFIELD ItemName ");
        builder.Write(", Quantity: ");
        builder.InsertField(" MERGEFIELD Quantity ");
        builder.InsertParagraph();

        // Insert the end of the mail merge region.
        // The region is closed with a MERGEFIELD that has the TableEnd: prefix.
        builder.InsertField(" MERGEFIELD TableEnd:OrderItems");

        // Prepare a DataTable that matches the region name and contains the data.
        DataTable orderItems = new DataTable("OrderItems");
        orderItems.Columns.Add("ItemName");
        orderItems.Columns.Add("Quantity");
        orderItems.Rows.Add(new object[] { "Apple", "10" });
        orderItems.Rows.Add(new object[] { "Banana", "5" });
        orderItems.Rows.Add(new object[] { "Orange", "8" });

        // Execute the mail merge using the DataTable. The region will repeat for each row.
        doc.MailMerge.ExecuteWithRegions(orderItems);

        // Save the resulting document to disk.
        doc.Save("OrderMergeResult.docx");
    }
}
