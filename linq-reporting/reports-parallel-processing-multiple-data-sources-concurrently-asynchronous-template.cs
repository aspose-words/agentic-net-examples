using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ParallelReportGenerator
{
    // Generates reports for multiple data sources concurrently.
    // Each report uses the same template but a different data source.
    public static async Task GenerateReportsAsync(
        string templatePath,
        IList<object> dataSources,
        IList<string> dataSourceNames,
        IList<string> outputPaths)
    {
        if (dataSources == null) throw new ArgumentNullException(nameof(dataSources));
        if (outputPaths == null) throw new ArgumentNullException(nameof(outputPaths));
        if (dataSources.Count != outputPaths.Count)
            throw new ArgumentException("The number of data sources must match the number of output paths.");

        // If explicit data source names are not provided, use nulls.
        var names = dataSourceNames ?? new List<string>(new string[dataSources.Count]);

        var tasks = new List<Task>();

        for (int i = 0; i < dataSources.Count; i++)
        {
            int index = i; // Capture loop variable for the lambda.
            tasks.Add(Task.Run(() =>
            {
                // Load the template document.
                Document template = new Document(templatePath);

                // Create a ReportingEngine instance.
                ReportingEngine engine = new ReportingEngine();

                // Build the report using the overload that accepts a data source name.
                // If the name is null or empty, the engine will treat it as omitted.
                engine.BuildReport(template, dataSources[index], names[index]);

                // Save the generated report.
                template.Save(outputPaths[index]);
            }));
        }

        // Await all report generation tasks.
        await Task.WhenAll(tasks);
    }

    // Example usage.
    public static async Task Main()
    {
        // Path to the Word template containing LINQ Reporting Engine tags.
        string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Example data sources – could be any supported type (e.g., DataSet, custom POCO, JsonDataSource, etc.).
        var dataSources = new List<object>
        {
            new { Title = "Quarter 1", Amount = 12345.67 },
            new { Title = "Quarter 2", Amount = 23456.78 },
            new { Title = "Quarter 3", Amount = 34567.89 }
        };

        // Optional names to reference the data sources inside the template.
        var dataSourceNames = new List<string> { "Q1", "Q2", "Q3" };

        // Output file paths for each generated report.
        var outputPaths = new List<string>
        {
            @"C:\Reports\Report_Q1.docx",
            @"C:\Reports\Report_Q2.docx",
            @"C:\Reports\Report_Q3.docx"
        };

        // Generate all reports in parallel.
        await GenerateReportsAsync(templatePath, dataSources, dataSourceNames, outputPaths);
    }
}
