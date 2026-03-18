using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

class MailMergeTableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table that will act as a placeholder for mail merge data.
        Table table = builder.StartTable();

        // Insert the first cell – this will contain the mail merge region start tag.
        builder.InsertCell();
        // Insert a MERGEFIELD that marks the beginning of the region.
        // The field name must be in the format "TableStart:<TableName>".
        builder.InsertField(" MERGEFIELD TableStart:Orders ");

        // Insert the second cell – this will hold the actual data field.
        builder.InsertCell();
        // Insert a MERGEFIELD for a column in the data source (e.g., "ProductName").
        builder.InsertField(" MERGEFIELD ProductName ");

        // Insert the third cell – this will contain the mail merge region end tag.
        builder.InsertCell();
        // Insert a MERGEFIELD that marks the end of the region.
        // The field name must be in the format "TableEnd:<TableName>".
        builder.InsertField(" MERGEFIELD TableEnd:Orders ");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // OPTIONAL: Demonstrate a mail merge with regions using a sample DataTable.
        // Create a DataTable that matches the region name ("Orders") and column name ("ProductName").
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("ProductName", typeof(string));
        orders.Rows.Add("Apple");
        orders.Rows.Add("Banana");
        orders.Rows.Add("Cherry");

        // Execute the mail merge. The table rows will be repeated for each record in the DataTable.
        doc.MailMerge.ExecuteWithRegions(orders);

        // Save the resulting document.
        doc.Save("MailMergeTablePlaceholder.docx");
    }
}
