using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the code page provider is registered (required for CSV parsing on some platforms).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the temporary files.
            string dataFilePath = "data.csv";
            string templateFilePath = "template.docx";
            string outputFilePath = "report.docx";

            // 1. Create a sample CSV file with a header and some rows.
            // Columns: Id, Name, Status
            string[] csvLines =
            {
                "Id,Name,Status",
                "1,Apple,Available",
                "2,Banana,OutOfStock",
                "3,Cherry,Available",
                "4,Date,Discontinued",
                "5,Elderberry,Available",
                "6,Fig,OutOfStock"
            };
            File.WriteAllLines(dataFilePath, csvLines, Encoding.UTF8);

            // 2. Build a Word template programmatically.
            // The template uses a LINQ GroupBy expression to summarize the data by Status.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Product Status Summary");
            builder.Writeln();

            // GroupBy expression: persons.GroupBy(p => p.Status)
            // 'persons' is the name we will give to the CSV data source.
            builder.Writeln("<<foreach [g in persons.GroupBy(p => p.Status)]>>");
            builder.Writeln("Status: <<[g.Key]>>   Count: <<[g.Count()]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before BuildReport according to the rules).
            templateDoc.Save(templateFilePath);

            // 3. Load the template back (demonstrates the load step).
            Document loadedTemplate = new Document(templateFilePath);

            // 4. Prepare the CSV data source with header support.
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            // Use default delimiter ','; other options can be set if needed.
            CsvDataSource csvDataSource = new CsvDataSource(dataFilePath, loadOptions);

            // 5. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are needed for this simple example.
            engine.Options = ReportBuildOptions.None;

            // The data source name "persons" must match the name used in the template tags.
            engine.BuildReport(loadedTemplate, csvDataSource, "persons");

            // 6. Save the generated report.
            loadedTemplate.Save(outputFilePath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputFilePath)}");
        }
    }
}
