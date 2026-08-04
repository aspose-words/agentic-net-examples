using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Duration { get; set; } = "01:02:03";
}

public class Program
{
    public static void Main()
    {
        // Create a template document with a LINQ Reporting tag that parses a duration string.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Parsed duration (hh:mm:ss): <<[TimeSpan.Parse(Duration)]>>");
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        Model model = new Model();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register System.TimeSpan to allow static method calls like TimeSpan.Parse in the template.
        engine.KnownTypes.Add(typeof(TimeSpan));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
