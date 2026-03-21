using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

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

        var names = dataSourceNames ?? new List<string>(new string[dataSources.Count]);

        var tasks = new List<Task>();

        for (int i = 0; i < dataSources.Count; i++)
        {
            int index = i; // Capture loop variable for the lambda.
            tasks.Add(Task.Run(async () =>
            {
                // Load the template (simple text file).
                string template = await File.ReadAllTextAsync(templatePath);

                // Prepare data for replacement.
                var data = dataSources[index];
                var name = names[index] ?? $"DataSource{index + 1}";

                // Very simple placeholder replacement using reflection.
                foreach (var prop in data.GetType().GetProperties())
                {
                    string placeholder = $"{{{{{prop.Name}}}}}";
                    string value = prop.GetValue(data)?.ToString() ?? string.Empty;
                    template = template.Replace(placeholder, value);
                }

                // Also replace a placeholder for the data source name if present.
                template = template.Replace("{{DataSourceName}}", name);

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPaths[index])!;
                Directory.CreateDirectory(outputDir);

                // Save the generated report.
                await File.WriteAllTextAsync(outputPaths[index], template);
            }));
        }

        await Task.WhenAll(tasks);
    }

    // Example usage.
    public static async Task Main(string[] args)
    {
        // Create a temporary folder for the demo.
        string baseDir = Path.Combine(Path.GetTempPath(), "ParallelReportDemo");
        Directory.CreateDirectory(baseDir);

        // Path to the simple text template containing placeholders.
        string templatePath = Path.Combine(baseDir, "ReportTemplate.txt");
        await File.WriteAllTextAsync(templatePath,
            "Report for {{DataSourceName}}\nTitle: {{Title}}\nAmount: {{Amount}}\nGenerated on {{Date}}");

        // Example data sources – could be any supported type (e.g., custom POCO, anonymous type, etc.).
        var dataSources = new List<object>
        {
            new { Title = "Quarter 1", Amount = 12345.67, Date = DateTime.Now.ToShortDateString() },
            new { Title = "Quarter 2", Amount = 23456.78, Date = DateTime.Now.ToShortDateString() },
            new { Title = "Quarter 3", Amount = 34567.89, Date = DateTime.Now.ToShortDateString() }
        };

        // Optional names to reference the data sources inside the template.
        var dataSourceNames = new List<string> { "Q1", "Q2", "Q3" };

        // Output file paths for each generated report.
        var outputPaths = new List<string>
        {
            Path.Combine(baseDir, "Report_Q1.txt"),
            Path.Combine(baseDir, "Report_Q2.txt"),
            Path.Combine(baseDir, "Report_Q3.txt")
        };

        // Generate all reports in parallel.
        await GenerateReportsAsync(templatePath, dataSources, dataSourceNames, outputPaths);

        Console.WriteLine("Reports generated in: " + baseDir);
    }
}
