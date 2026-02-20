using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Prepare a sample DataTable.
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("Item");
        dataTable.Columns.Add("Quantity", typeof(int));
        dataTable.Rows.Add("Apples", 20);
        dataTable.Rows.Add("Bananas", 40);
        dataTable.Rows.Add("Carrots", 50);

        // Create a new Word document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Add a header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Iterate through the DataTable rows and add them to the Word table.
        foreach (DataRow dr in dataTable.Rows)
        {
            builder.InsertCell();
            builder.Write(dr["Item"].ToString());

            builder.InsertCell();
            builder.Write(dr["Quantity"].ToString());

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Adjust table layout to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("TableFromDataTable.docx");
    }
}
