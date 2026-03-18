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
        //    The region is defined by the special MERGEFIELDs TableStart:Items and TableEnd:Items.
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

        // Start of the foreach band (mail‑merge region).
        builder.InsertCell();
        builder.InsertField("MERGEFIELD TableStart:Items");
        builder.EndRow();

        // Data row – each column contains a MERGEFIELD.
        // Column 1 – Item name.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Item");
        // Column 2 – Quantity.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity");
        // Column 3 – Unit price.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD UnitPrice");
        // Column 4 – Calculated line total.
        // The expression field uses nested MERGEFIELDs to reference the current iteration's values.
        // The numeric format \# "$#,##0.00" formats the result as currency.
        builder.InsertCell();
        builder.InsertField("= { MERGEFIELD Quantity } * { MERGEFIELD UnitPrice } \\# \"$#,##0.00\"");
        builder.EndRow();

        // End of the foreach band.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD TableEnd:Items");
        builder.EndRow();

        builder.EndTable();

        // 3. Prepare a data source for the mail merge.
        //    The DataTable must contain the columns referenced by the MERGEFIELDs.
        DataTable table = new DataTable("Items");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("UnitPrice", typeof(decimal));

        table.Rows.Add("Pen", 10, 1.20m);
        table.Rows.Add("Notebook", 5, 3.50m);
        table.Rows.Add("Eraser", 12, 0.75m);

        // 4. Execute the mail merge.
        //    The region name ("Items") matches the name used in TableStart/TableEnd fields.
        doc.MailMerge.Execute(table);

        // 5. Update all fields so that the calculated totals are evaluated.
        doc.UpdateFields();

        // 6. Save the resulting document.
        doc.Save("LineTotal_ForeachBand.docx");
    }
}
