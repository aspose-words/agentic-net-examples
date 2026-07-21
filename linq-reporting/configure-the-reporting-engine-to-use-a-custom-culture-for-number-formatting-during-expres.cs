using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Set a custom culture (German) for number formatting.
        CultureInfo customCulture = new CultureInfo("de-DE");
        System.Threading.Thread.CurrentThread.CurrentCulture = customCulture;

        // Sample data model.
        var model = new ReportModel { Price = 1234.56 };

        // Create a template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Price: <<[model.Price]>>");

        // Save the template (demonstrates the create‑save lifecycle).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template (demonstrates the load step).
        Document loadedTemplate = new Document(templatePath);

        // Build the report using the custom culture.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save("Report.docx");
    }

    // Public data model with a numeric property.
    public class ReportModel
    {
        public double Price { get; set; } = 0;
    }
}
