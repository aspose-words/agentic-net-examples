using System;
using System.Data;
using System.IO;
using Aspose.Words;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert MERGEFIELDs for CustomerName and Address.
            builder.InsertField(" MERGEFIELD CustomerName ");
            builder.Writeln(); // Move to next line.
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare a data source with two records.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
            table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Execute the mail merge using the DataTable.
            doc.MailMerge.Execute(table);

            // Save the resulting document to the executable's folder.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MailMergeResult.docx");
            doc.Save(outputPath);
        }
    }
}
