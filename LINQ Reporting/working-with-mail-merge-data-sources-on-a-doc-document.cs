using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    class Program
    {
        static void Main()
        {
            // Define a folder where the resulting document will be saved.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(artifactsDir);

            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add MERGEFIELDs that will be filled by the mail merge.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare a DataTable that matches the field names in the document.
            DataTable table = new DataTable("Customers");
            table.Columns.Add("FullName");
            table.Columns.Add("Address");

            // Add sample rows.
            table.Rows.Add("Thomas Hardy", "120 Hanover Sq., London");
            table.Rows.Add("Paolo Accorti", "Via Monte Bianco 34, Torino");

            // Perform the mail merge using the DataTable as the data source.
            doc.MailMerge.Execute(table);

            // Save the merged document.
            string outputPath = Path.Combine(artifactsDir, "MailMergeFromDataTable.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Mail merge completed. Document saved to: {outputPath}");
        }
    }
}
