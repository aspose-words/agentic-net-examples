using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document that contains a static DateTime call.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        // Correct tag syntax for static member access uses a dot.
        builder.Writeln("Current date and time: <<[DateTime.Now]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine and add System.DateTime to known types.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report using an empty model.
        var model = new Model();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }

    // Empty model class required by BuildReport.
    public class Model
    {
    }
}
