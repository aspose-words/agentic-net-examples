using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportMemoryMeasurement
{
    static void Main()
    {
        // Path to the template document that contains reporting tags.
        string templatePath = "Template.docx";          // TODO: replace with actual template file path
        // Path where the generated report will be saved.
        string outputPath = "Report.docx";              // TODO: replace with desired output file path

        // Load the template document.
        Document template = new Document(templatePath);

        // Create the reporting engine and enable the RemoveEmptyParagraphs option.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Prepare the data source for the report.
        // Replace this with the actual data source (e.g., a DataSet, an object, etc.).
        object dataSource = GetReportData(); // TODO: implement GetReportData()

        // Force a garbage collection to get a clean baseline memory measurement.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Record memory usage before building the report.
        long memoryBefore = Process.GetCurrentProcess().PrivateMemorySize64;

        // Build the report.
        engine.BuildReport(template, dataSource);

        // Record memory usage after building the report.
        long memoryAfter = Process.GetCurrentProcess().PrivateMemorySize64;

        // Calculate the memory consumed by the report generation.
        long memoryConsumed = memoryAfter - memoryBefore;

        // Output the memory consumption information.
        Console.WriteLine($"Memory before report generation: {memoryBefore:N0} bytes");
        Console.WriteLine($"Memory after report generation : {memoryAfter:N0} bytes");
        Console.WriteLine($"Memory consumed by report generation (RemoveEmptyParagraphs enabled): {memoryConsumed:N0} bytes");

        // Save the generated report.
        template.Save(outputPath);
    }

    // Placeholder method – replace with actual logic to obtain the data source.
    private static object GetReportData()
    {
        // Example: return a DataSet, a custom object, etc.
        // For demonstration purposes we return an empty anonymous object.
        return new { };
    }
}
