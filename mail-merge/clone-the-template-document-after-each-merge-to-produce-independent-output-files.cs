using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a template document with merge fields.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(": ");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Prepare a data source with several records.
            DataTable data = new DataTable("Recipients");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Message");

            data.Rows.Add("John", "Doe", "Hello! This is the first message.");
            data.Rows.Add("Jane", "Smith", "Greetings from the second record.");
            data.Rows.Add("Bob", "Johnson", "Third message goes here.");

            // Perform a separate mail merge for each record, cloning the template each time.
            for (int i = 0; i < data.Rows.Count; i++)
            {
                // Clone the template to obtain an independent document.
                Document mergedDoc = (Document)template.Clone(true);

                // Execute mail merge for the current row.
                mergedDoc.MailMerge.Execute(data.Rows[i]);

                // Save the merged document to a distinct file.
                string fileName = $"MergedDocument_{i + 1}.docx";
                mergedDoc.Save(fileName, SaveFormat.Docx);
            }
        }
    }
}
