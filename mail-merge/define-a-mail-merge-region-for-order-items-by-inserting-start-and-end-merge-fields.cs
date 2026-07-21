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

        // Add a title.
        builder.Writeln("Order Summary:");

        // Insert the start of the mail merge region named "OrderItems".
        builder.InsertField(" MERGEFIELD TableStart:OrderItems");

        // Inside the region insert fields for each column.
        builder.Write("Item: ");
        builder.InsertField(" MERGEFIELD Item");
        builder.Write(", Quantity: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.InsertParagraph();

        // Insert the end of the region.
        builder.InsertField(" MERGEFIELD TableEnd:OrderItems");

        // Prepare a DataTable that matches the region name and column names.
        DataTable orderTable = new DataTable("OrderItems");
        orderTable.Columns.Add("Item");
        orderTable.Columns.Add("Quantity");
        orderTable.Rows.Add(new object[] { "Apple", "10" });
        orderTable.Rows.Add(new object[] { "Banana", "5" });
        orderTable.Rows.Add(new object[] { "Orange", "8" });

        // Execute the mail merge with regions using the DataTable.
        doc.MailMerge.ExecuteWithRegions(orderTable);

        // Save the resulting document.
        doc.Save("OrderReport.docx");
    }
}
