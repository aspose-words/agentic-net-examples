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
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocs");
            Directory.CreateDirectory(outputDir);

            // 1. Create a mail‑merge template document with three fields.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // 2. Build a data source containing several records.
            DataTable data = new DataTable("Recipients");
            data.Columns.Add("FirstName");
            data.Columns.Add("LastName");
            data.Columns.Add("Message");
            data.Rows.Add("John", "Doe", "Hello! This is the first message.");
            data.Rows.Add("Jane", "Smith", "Greetings from the second recipient.");
            data.Rows.Add("Bob", "Johnson", "A third message for you.");

            // 3. For each record clone the template, perform a mail merge, and save the result.
            int index = 1;
            foreach (DataRow row in data.Rows)
            {
                // Clone the original template so each merge starts from a clean document.
                Document mergedDoc = (Document)template.Clone(true);

                // Execute mail merge for the current row only.
                mergedDoc.MailMerge.Execute(row);

                // Save the merged document with a unique name.
                string fileName = Path.Combine(outputDir, $"MergedDocument_{index}.docx");
                mergedDoc.Save(fileName);
                index++;
            }

            // All merged documents are now stored in the "MergedDocs" folder.
        }
    }
}
