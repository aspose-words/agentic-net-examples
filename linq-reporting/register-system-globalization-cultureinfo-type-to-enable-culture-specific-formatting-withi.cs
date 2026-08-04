using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public DateTime Date { get; set; } = DateTime.Now;
    public double Number { get; set; } = 12345.6789;
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Tag that formats a date using a specific culture (French).
        builder.Writeln("Date (fr-FR): <<[model.Date.ToString(\"D\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");

        // Tag that formats a number using a specific culture (German).
        builder.Writeln("Number (de-DE): <<[model.Number.ToString(\"N\", CultureInfo.GetCultureInfo(\"de-DE\"))]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Create the reporting engine and register CultureInfo type.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(CultureInfo));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
