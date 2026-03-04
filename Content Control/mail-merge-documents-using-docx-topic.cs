using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document (creation rule).
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Insert merge fields that will be replaced during the mail merge.
            builder.InsertField(" MERGEFIELD CustomerName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare a DataTable that contains the data for the mail merge.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
            table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

            // Execute the mail merge using the DataTable as the data source.
            sourceDoc.MailMerge.Execute(table);

            // Save the merged document to a DOCX file (save rule).
            sourceDoc.Save("MergedDocument.docx", SaveFormat.Docx);
        }
    }
}
