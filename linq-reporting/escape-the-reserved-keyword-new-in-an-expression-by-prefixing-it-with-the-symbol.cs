using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Property name "new" is a C# reserved keyword, escaped with @.
    public string @new { get; set; } = "EscapedKeyword";
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with a LINQ Reporting tag.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        // The tag references the property named "new" in the model.
        builder.Writeln("Value: <<[model.new]>>");
        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template document for reporting.
        var doc = new Document(templatePath);

        // Step 3: Build the report using the model that contains the escaped property.
        var engine = new ReportingEngine();
        var model = new ReportModel(); // model.new will be used in the template.
        engine.BuildReport(doc, model, "model");

        // Step 4: Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
