using System;
using System.Data;
using Aspose.Words;

namespace AsposeWordsMailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert MERGEFIELDs that will be populated by the mail merge.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD CustomerName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare a DataTable that contains the data for the mail merge.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
            table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

            // Execute the mail merge using the whole table – one document will be generated per row.
            doc.MailMerge.Execute(table);

            // Save the merged document to a DOCX file.
            doc.Save("MailMergeResult.docx");
        }
    }
}
