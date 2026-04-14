using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a mail‑merge source document containing the required fields.
            Document doc = CreateSourceDocument();

            // Build a data table that matches the merge field names.
            DataTable data = CreateDataTable();

            // Perform the mail merge using the data table.
            doc.MailMerge.Execute(data);

            // Save the merged result as a DOCX file.
            doc.Save("MergedDocument.docx", SaveFormat.Docx);
        }

        // Creates a blank document and inserts MERGEFIELDs.
        private static Document CreateSourceDocument()
        {
            Document document = new Document();
            DocumentBuilder builder = new DocumentBuilder(document);

            builder.InsertField("MERGEFIELD CustomerName");
            builder.InsertParagraph();
            builder.InsertField("MERGEFIELD Address");

            return document;
        }

        // Generates a DataTable with columns that correspond to the merge fields.
        private static DataTable CreateDataTable()
        {
            DataTable table = new DataTable("Customers");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");

            table.Rows.Add("John Doe", "123 Main St, Anytown");
            table.Rows.Add("Jane Smith", "456 Oak Ave, Othertown");

            return table;
        }
    }
}
