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
            // Create a template document that contains the merge fields.
            Document template = CreateTemplate();

            // Build a data table with several records.
            DataTable data = CreateDataTable();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Iterate over each record, perform a mail merge, and save the result as a PDF.
            int recordIndex = 1;
            foreach (DataRow row in data.Rows)
            {
                // Clone the template for the current record.
                Document doc = (Document)template.Clone();

                // Merge the current row into the document.
                doc.MailMerge.Execute(row);

                // Save the merged document as a PDF file.
                string pdfPath = Path.Combine(outputDir, $"MergedDocument_{recordIndex}.pdf");
                doc.Save(pdfPath, SaveFormat.Pdf);

                recordIndex++;
            }
        }

        // Creates a simple Word document with MERGEFIELDs.
        private static Document CreateTemplate()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            return doc;
        }

        // Generates a DataTable containing sample data for the mail merge.
        private static DataTable CreateDataTable()
        {
            DataTable table = new DataTable("Recipients");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Columns.Add("Message");

            table.Rows.Add("John", "Doe", "Hello! This is the first message.");
            table.Rows.Add("Jane", "Smith", "Greetings from Aspose.Words.");
            table.Rows.Add("Bob", "Johnson", "Your order has been shipped.");

            return table;
        }
    }
}
