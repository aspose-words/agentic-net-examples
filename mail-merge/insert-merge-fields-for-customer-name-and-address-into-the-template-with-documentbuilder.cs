using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

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
            builder.InsertParagraph(); // Move to a new paragraph.
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare a data source with one record.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");

            // Execute the mail merge using the data table.
            doc.MailMerge.Execute(table);

            // Save the merged document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedDocument.docx");
            doc.Save(outputPath);
        }
    }
}
