using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create an output folder relative to the executable location.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Build a template document that contains the merge fields.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Save the template for reference.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare a data source (DataTable) with several records.
            // -----------------------------------------------------------------
            DataTable table = new DataTable("Recipients");
            table.Columns.Add("FirstName");
            table.Columns.Add("LastName");
            table.Columns.Add("Message");

            table.Rows.Add("John", "Doe", "Hello! This is a merged message.");
            table.Rows.Add("Jane", "Smith", "Welcome to Aspose.Words mail merge.");
            table.Rows.Add("Bob", "Johnson", "Your order has been shipped.");

            // -----------------------------------------------------------------
            // 3. Merge all records into a single document.
            // -----------------------------------------------------------------
            Document mergedAll = (Document)template.Clone();
            mergedAll.MailMerge.Execute(table);
            string mergedAllPath = Path.Combine(outputDir, "MergedAll.docx");
            mergedAll.Save(mergedAllPath);

            // -----------------------------------------------------------------
            // 4. Clone the template for each record and generate separate files.
            // -----------------------------------------------------------------
            int index = 1;
            foreach (DataRow row in table.Rows)
            {
                Document individual = (Document)template.Clone();
                individual.MailMerge.Execute(row);

                string fileName = $"Merged_{index}_{row["FirstName"]}_{row["LastName"]}.docx";
                string filePath = Path.Combine(outputDir, fileName);
                individual.Save(filePath);
                index++;
            }
        }
    }
}
