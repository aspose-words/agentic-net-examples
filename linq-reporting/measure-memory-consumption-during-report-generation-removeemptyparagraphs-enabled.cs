using System;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    public string Name { get; set; }
    public int Age { get; set; }
}

class ReportMemoryMeasurement
{
    static void Main()
    {
        // Create a simple template document with reporting tags.
        Document template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[ReportData.Name]>>");
        builder.Writeln("Age: <<[ReportData.Age]>>");

        // Path where the generated report will be saved.
        string outputPath = "Report.docx";

        // Create the reporting engine and enable the RemoveEmptyParagraphs option.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Prepare a visible data source for the report.
        var dataSource = new ReportData
        {
            Name = "John Doe",
            Age = 30
        };

        // Force a garbage collection to get a clean baseline memory measurement.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Record memory usage before building the report.
        long memoryBefore = Process.GetCurrentProcess().PrivateMemorySize64;

        // Build the report using the data source name "ReportData".
        engine.BuildReport(template, dataSource, "ReportData");

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
}
