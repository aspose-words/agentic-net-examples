using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table that will contain the mail merge region.
        builder.StartTable();

        // Insert the first cell and place the TableStart field for the region named "Products".
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableStart:Products ");

        // Insert a cell for the first column and add a merge field for the product name.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD ProductName ");

        // Insert a cell for the second column and add a merge field for the price.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Price ");

        // Insert the last cell and place the TableEnd field to close the region.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableEnd:Products ");

        // End the row that will be repeated for each record.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Prepare a DataTable that matches the region name and contains the data.
        DataTable products = new DataTable("Products");
        products.Columns.Add("ProductName");
        products.Columns.Add("Price");
        products.Rows.Add("Apple", "$1.00");
        products.Rows.Add("Banana", "$0.50");
        products.Rows.Add("Cherry", "$2.00");

        // Execute the mail merge with regions. The row inside the table will be duplicated for each record.
        doc.MailMerge.ExecuteWithRegions(products);

        // Save the resulting document.
        doc.Save("MailMergeTableRegion.docx");
    }
}
