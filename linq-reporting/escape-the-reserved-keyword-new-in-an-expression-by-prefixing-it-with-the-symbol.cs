using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Property name is the reserved keyword 'new', escaped with @ in C#.
    // The actual property name is "new", which the LINQ Reporting engine accesses without the @ prefix.
    public string @new { get; set; } = "Escaped keyword value";

    public string Name { get; set; } = "Sample Name";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a simple paragraph with a LINQ Reporting tag that accesses the escaped property.
        // The property name is referenced without the @ prefix in the template expression.
        builder.Writeln("Value of the escaped property 'new': <<[model.new]>>");
        builder.Writeln("Another field (Name): <<[model.Name]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data source
        ReportModel model = new ReportModel();

        // Create the reporting engine and generate the report
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the root object name "model"
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(reportPath);
    }
}
