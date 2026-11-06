using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a LINQ Reporting tag that references a property named "new".
        Document template = new();
        DocumentBuilder builder = new(template);
        // The property name is "new". In C# the identifier is written as @new, but the runtime name is "new",
        // so the expression must use the plain name without the @ prefix.
        builder.Writeln("Escaped keyword example: <<[model.new]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Prepare the data model. The property name is the reserved keyword "new",
        // escaped with the @ symbol in C# source code.
        var model = new SampleModel
        {
            @new = "Reserved keyword value"
        };

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Data model with a property named "new" escaped using @.
public class SampleModel
{
    // The identifier is written as @new, but the actual property name is "new".
    public string @new { get; set; } = string.Empty;
}
