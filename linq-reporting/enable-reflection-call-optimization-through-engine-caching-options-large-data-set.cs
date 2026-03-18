using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReflectionOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Enable reflection call optimization for the ReportingEngine.
            // This reduces the overhead of reflection when processing large data sets.
            ReportingEngine.UseReflectionOptimization = true;

            // Path to the template document that contains LINQ Reporting Engine tags.
            // Example template content: <<[Customers.Name]>>
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a large data set (e.g., 20,000 rows) to be merged.
            DataSet data = CreateLargeDataSet(20000);

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the large data set.
            engine.BuildReport(doc, data, string.Empty);

            // Configure save options to lower memory consumption during saving.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
            saveOptions.MemoryOptimization = true; // Optimize memory for large documents.

            // Save the generated report.
            string outputPath = @"C:\Docs\Report_Output.docx";
            doc.Save(outputPath, saveOptions);
        }

        // Helper method that creates a DataSet with a single DataTable containing a large number of rows.
        private static DataSet CreateLargeDataSet(int rowCount)
        {
            DataTable table = new DataTable("Customers");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("Phone", typeof(string));

            for (int i = 1; i <= rowCount; i++)
            {
                table.Rows.Add(
                    $"Customer {i}",
                    $"customer{i}@example.com",
                    $"+1-555-010{i:D4}"
                );
            }

            DataSet ds = new DataSet();
            ds.Tables.Add(table);
            return ds;
        }
    }
}
