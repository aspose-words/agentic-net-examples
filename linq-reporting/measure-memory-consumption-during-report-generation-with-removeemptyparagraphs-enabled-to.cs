using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Person
{
    public string Name { get; set; } = "";
    public int? Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
    public string EmptyTag { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = null },   // Age will be empty.
                new Person { Name = "", Age = 25 }        // Name will be empty.
            },
            EmptyTag = "" // Resolves to an empty string.
        };

        // Create the template document.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Paragraph that becomes empty after processing.
        builder.Writeln("<<[EmptyTag]>>");

        // Begin the foreach block.
        builder.Writeln("<<foreach [p in Persons]>>");

        // Table with header and data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row (repeated for each person).
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Age]>>");
        builder.EndRow();

        // End of table.
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Measure memory before building the report.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long memoryBefore = Process.GetCurrentProcess().PrivateMemorySize64;

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Measure memory after building the report.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long memoryAfter = Process.GetCurrentProcess().PrivateMemorySize64;

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Output memory consumption.
        Console.WriteLine($"Memory before report generation: {memoryBefore / 1024} KB");
        Console.WriteLine($"Memory after  report generation: {memoryAfter / 1024} KB");
        Console.WriteLine($"Memory increase: {(memoryAfter - memoryBefore) / 1024} KB");
    }
}
