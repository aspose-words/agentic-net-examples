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
            // Create a template document in memory.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert mail‑merge region tags and a field.
            builder.Writeln("<<TableStart:Employees>>");
            builder.Writeln("Name: <<Name>>");
            builder.Writeln("<<TableEnd:Employees>>");

            // Prepare data.
            DataSet dataSet = new DataSet();
            DataTable table = new DataTable("Employees");
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("John Doe");
            table.Rows.Add("Jane Smith");
            dataSet.Tables.Add(table);

            // Execute mail merge with regions.
            template.MailMerge.ExecuteWithRegions(dataSet);

            // Save the result to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmployeeReport.docx");
            template.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");
        }
    }
}
