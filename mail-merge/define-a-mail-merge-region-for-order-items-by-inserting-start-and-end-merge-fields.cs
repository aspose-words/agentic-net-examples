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

        // Insert the start tag of the mail merge region.
        // The region name "OrderItems" must match the DataTable name used later.
        builder.InsertField(" MERGEFIELD TableStart:OrderItems");

        // Insert fields that will be populated for each row of the region.
        builder.Write("Item: ");
        builder.InsertField(" MERGEFIELD ItemName ");
        builder.Write(", Qty: ");
        builder.InsertField(" MERGEFIELD Quantity ");
        builder.InsertParagraph(); // Separate each record with a paragraph.

        // Insert the end tag of the mail merge region.
        builder.InsertField(" MERGEFIELD TableEnd:OrderItems");

        // Prepare a DataTable that matches the region name and contains the data.
        DataTable orderItems = new DataTable("OrderItems");
        orderItems.Columns.Add("ItemName", typeof(string));
        orderItems.Columns.Add("Quantity", typeof(int));

        // Add some sample rows.
        orderItems.Rows.Add("Rugby Ball", 2);
        orderItems.Rows.Add("Team Jersey", 5);
        orderItems.Rows.Add("Water Bottle", 3);

        // Execute the mail merge using the region defined above.
        doc.MailMerge.ExecuteWithRegions(orderItems);

        // Save the resulting document.
        doc.Save("MailMergeRegionExample.docx");
    }
}
