using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // File names in the current working directory.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // 1. Create a Word template with a LINQ Reporting tag that accesses a property named "new".
        //    The property is declared as @new in C#, but the template refers to it without the @ prefix.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("<<[model.new]>>"); // The tag will be replaced by the property value.
        doc.Save(templatePath);

        // 2. Prepare the data model. The property name is a reserved keyword,
        //    so it is declared as '@new' in C#.
        var model = new ReportModel
        {
            @new = "Escaped keyword value"
        };

        // 3. Load the template and build the report using ReportingEngine.
        var template = new Document(templatePath);
        var engine = new ReportingEngine();
        // The root object name used in the template is "model".
        engine.BuildReport(template, model, "model");

        // 4. Save the generated report.
        template.Save(outputPath);
    }
}

// Data model class with a property named "new" escaped as '@new'.
public class ReportModel
{
    // Property name is a reserved keyword; the '@' prefix allows it.
    public string @new { get; set; } = string.Empty;
}
