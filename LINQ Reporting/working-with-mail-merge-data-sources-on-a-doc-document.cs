using System;
using System.Data;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert MERGEFIELDs that will be populated by the mail merge.
        builder.InsertField(" MERGEFIELD CustomerName ");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD Address ");

        // Build a DataTable that will serve as the mail merge data source.
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName");
        table.Columns.Add("Address");
        table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
        table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

        // Perform the mail merge using the DataTable.
        doc.MailMerge.Execute(table);

        // Save the merged document.
        doc.Save("MailMergeResult.docx");
    }
}
