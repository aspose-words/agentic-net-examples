using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting tag that accesses System.Math static members.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Value of PI: <<[Math.PI]>>");
        builder.Writeln("Sin(0) = <<[Math.Sin(0)]>>");
        // Save the template locally.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Configure the ReportingEngine and expose System.Math to the template.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(System.Math));

        // Build the report. No data source is required because only static members are used.
        engine.BuildReport(report, new object(), "");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
