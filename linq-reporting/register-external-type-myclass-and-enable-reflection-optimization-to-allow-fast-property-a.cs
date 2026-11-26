using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MyClass
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build template document with LINQ Reporting tags
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello, <<[my.Name]>>!");
        builder.Writeln("Value: <<[my.Value]>>");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load template for report generation
        Document report = new Document(templatePath);

        // Sample data
        MyClass data = new MyClass
        {
            Name = "World",
            Value = 42
        };

        // Enable reflection optimization
        ReportingEngine.UseReflectionOptimization = true;

        // Build report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, data, "my");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        report.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
