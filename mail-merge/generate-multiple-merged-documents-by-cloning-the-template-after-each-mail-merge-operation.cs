using System;
using System.Data;
using Aspose.Words;

namespace MailMergeCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a mail merge template document with three merge fields.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Prepare a data source containing several records.
            DataTable data = new DataTable("Recipients");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Message");
            data.Rows.Add("John", "Doe", "Hello! This is the first message.");
            data.Rows.Add("Jane", "Smith", "Greetings from the second recipient.");
            data.Rows.Add("Bob", "Johnson", "Your third message is here.");

            // Generate a separate merged document for each row by cloning the template.
            int index = 1;
            foreach (DataRow row in data.Rows)
            {
                // Clone the template to obtain a fresh document for the current record.
                Document mergedDoc = (Document)template.Clone(true);

                // Perform mail merge using the current DataRow.
                mergedDoc.MailMerge.Execute(row);

                // Save the merged document to a file.
                string fileName = $"MergedDocument_{index}.docx";
                mergedDoc.Save(fileName);
                index++;
            }
        }
    }
}
