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

        // Insert the start of the mail merge region named "Orders".
        builder.InsertField(" MERGEFIELD TableStart:Orders");

        // Fields inside the region that will be filled from the data source.
        builder.Write("Item: ");
        builder.InsertField(" MERGEFIELD ItemName");
        builder.Write("\tQuantity: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.InsertParagraph();

        // Insert the end of the mail merge region.
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // Create a DataTable that matches the region name.
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("ItemName");
        orders.Columns.Add("Quantity");
        orders.Rows.Add(new object[] { "Rugby Ball", 2 });
        orders.Rows.Add(new object[] { "Soccer Jersey", 1 });
        orders.Rows.Add(new object[] { "Baseball Cap", 3 });

        // Execute the mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(orders);

        // Save the merged document.
        doc.Save("OrderMergeResult.docx");
    }
}
