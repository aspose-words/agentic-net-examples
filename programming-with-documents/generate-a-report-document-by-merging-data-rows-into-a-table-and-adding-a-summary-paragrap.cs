using System;
using System.Data;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header row.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Row that will be repeated by mail merge.
        // Insert TableStart field to mark the beginning of the region.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableStart:Products ");

        // Data fields.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Product ");
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Quantity ");
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Price ");

        // Insert TableEnd field to mark the end of the region.
        builder.InsertField(" MERGEFIELD TableEnd:Products ");
        builder.EndRow();

        builder.EndTable();

        // Prepare sample data.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Product");
        table.Columns.Add("Quantity");
        table.Columns.Add("Price");

        table.Rows.Add("Apples", "10", "$1.20");
        table.Rows.Add("Bananas", "5", "$0.80");
        table.Rows.Add("Carrots", "7", "$0.60");

        // Execute mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(table);

        // Add a summary paragraph after the table.
        DocumentBuilder summaryBuilder = new DocumentBuilder(doc);
        summaryBuilder.MoveToDocumentEnd();
        summaryBuilder.Writeln();
        summaryBuilder.Writeln($"Report generated with {table.Rows.Count} records.");

        // Save the resulting document.
        doc.Save("Report.docx");
    }
}
