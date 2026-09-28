using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new document and add merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<FirstName>> <<LastName>>,");
        builder.Writeln("Your order <<OrderNumber>> has been shipped.");

        // Prepare a data source for mail merge.
        DataTable table = new DataTable("Customers");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("OrderNumber");
        table.Rows.Add("John", "Doe", "12345");
        table.Rows.Add("Jane", "Smith", "67890");

        // Execute mail merge.
        doc.MailMerge.Execute(table);

        // Save the merged document as PDF.
        doc.Save("MergedOutput.pdf", SaveFormat.Pdf);
    }
}
