using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD for the customer's name.
        builder.InsertField(" MERGEFIELD CustomerName ");

        // Insert a line break.
        builder.Writeln();

        // Insert a MERGEFIELD for the customer's address.
        builder.InsertField(" MERGEFIELD Address ");

        // Prepare a simple data source with one record.
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName");
        table.Columns.Add("Address");
        table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");

        // Perform the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result to a file in the current directory.
        doc.Save("MergedDocument.docx");
    }
}
