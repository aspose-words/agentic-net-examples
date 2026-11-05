using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // Create a Word document that will serve as the LINQ Reporting template.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a tag that calls the static DateTime.Now property.
        builder.Writeln("Current time: <<[DateTime.Now]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Add System.DateTime to the list of known types so the template can use it.
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report. No root data object is needed for this example.
        engine.BuildReport(reportDoc, new object(), "model");

        // Save the generated report.
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine("Report generated successfully.");
    }
}
