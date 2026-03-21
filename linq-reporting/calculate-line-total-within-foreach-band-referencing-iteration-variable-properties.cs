using System;
using System.Data;
using Aspose.Words;

class MailMergeCalculateLineTotal
{
    static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Insert a table that will contain the mail‑merge region.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Unit Price");
        builder.InsertCell();
        builder.Write("Line Total");
        builder.EndRow();

        // Row that defines the mail‑merge region and contains the data fields.
        // TableStart and TableEnd must be in the same row.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD TableStart:Items");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Item");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD UnitPrice");
        builder.InsertCell();
        builder.InsertField("= { MERGEFIELD Quantity } * { MERGEFIELD UnitPrice } \\# \"$#,##0.00\"");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD TableEnd:Items");
        builder.EndRow();

        builder.EndTable();

        // 3. Prepare a data source for the mail merge.
        DataTable table = new DataTable("Items");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("UnitPrice", typeof(decimal));

        table.Rows.Add("Pen", 10, 1.20m);
        table.Rows.Add("Notebook", 5, 3.50m);
        table.Rows.Add("Eraser", 12, 0.75m);

        // 4. Execute the mail merge.
        doc.MailMerge.Execute(table);

        // 5. Update all fields so that the calculated totals are evaluated.
        doc.UpdateFields();

        // 6. Save the resulting document.
        doc.Save("LineTotal_ForeachBand.docx");
    }
}
