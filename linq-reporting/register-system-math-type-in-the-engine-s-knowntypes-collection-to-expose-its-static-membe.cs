using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use static members of System.Math in the template.
        builder.Writeln("The value of PI is: <<[Math.PI]>>");
        builder.Writeln("Square root of 16 is: <<[Math.Sqrt(16)]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Register System.Math so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(System.Math));

        // No data source is required for this example; an empty object is sufficient.
        engine.BuildReport(reportDoc, new object());

        // Save the generated report.
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
