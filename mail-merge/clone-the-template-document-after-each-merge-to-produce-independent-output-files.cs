using System;
using System.Data;
using System.IO;
using Aspose.Words;

namespace MailMergeCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a simple template document with merge fields.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Create a data source with several records.
            DataTable data = new DataTable("Customers");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Message");
            data.Rows.Add("John", "Doe", "Hello! This is the first message.");
            data.Rows.Add("Jane", "Smith", "Greetings from the second record.");
            data.Rows.Add("Bob", "Johnson", "Third message goes here.");

            // Ensure an output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Perform a separate mail merge for each record, cloning the template each time.
            int index = 1;
            foreach (DataRow row in data.Rows)
            {
                // Deep clone the template so each output file is independent.
                Document doc = (Document)template.Clone(true);

                // Execute mail merge for the current record.
                doc.MailMerge.Execute(row);

                // Save the merged document.
                string fileName = Path.Combine(outputDir, $"MergedDocument_{index}.docx");
                doc.Save(fileName);
                index++;
            }
        }
    }
}
