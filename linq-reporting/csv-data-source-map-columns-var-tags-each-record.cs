using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CsvReportingExample
{
    class Program
    {
        static void Main()
        {
            // Path to the CSV file. The first line contains column headers.
            string csvPath = @"C:\Data\People.csv";

            // Path to the template document that contains <<[data.ColumnName]>> tags.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\PeopleReport.docx";

            // -----------------------------------------------------------------
            // 1. Configure CSV loading options.
            //    - HasHeaders = true tells Aspose.Words that the first row holds column names.
            //    - Delimiter can be changed if the CSV uses a different separator.
            // -----------------------------------------------------------------
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
            {
                Delimiter = ',',      // default is ',', change if needed (e.g., ';')
                CommentChar = '#',    // default comment character
                QuoteChar = '"'       // default quote character
            };

            // -----------------------------------------------------------------
            // 2. Create the CSV data source using the file path and the options.
            // -----------------------------------------------------------------
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // -----------------------------------------------------------------
            // 3. Load the template document that contains the <<[data.ColumnName]>>
            //    tags. In the template each column from the CSV can be accessed
            //    via the data source name ("data" in this example).
            // -----------------------------------------------------------------
            Document template = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Build the report. The third argument is the name by which the
            //    data source will be referenced inside the template.
            //    Example tag in the template: <<[data.FirstName]>>.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            bool success = engine.BuildReport(template, csvDataSource, "data");

            if (!success)
            {
                Console.WriteLine("The template could not be parsed correctly.");
                return;
            }

            // -----------------------------------------------------------------
            // 5. Save the populated document.
            // -----------------------------------------------------------------
            template.Save(outputPath);
            Console.WriteLine($"Report generated successfully at: {outputPath}");
        }
    }
}
