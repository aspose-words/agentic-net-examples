using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

class Program
{
    static void Main()
    {
        // Create a simple blank document (no external template required).
        Document template = new Document();

        // Create a reporting engine instance.
        ReportingEngine engine = new ReportingEngine();

        // Register the Regex type so that its static members can be used inside template expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report using a visible data source type.
        var dataSource = new ReportData();
        engine.BuildReport(template, dataSource);

        // Save the generated report.
        template.Save("ReportOutput.docx");
    }
}
