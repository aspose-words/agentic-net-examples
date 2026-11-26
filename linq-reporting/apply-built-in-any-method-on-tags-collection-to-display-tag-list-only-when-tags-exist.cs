using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public List<string> Tags { get; set; } = new();
    public string Title { get; set; } = "Sample Report";
}

public class Program
{
    public static void Main()
    {
        // Create template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("<<[Title]>>");
        builder.Writeln();

        // Conditional display of tags using Any().
        builder.Writeln("<<if [Tags.Any()]>>");
        builder.Writeln("Tags:");
        builder.Writeln("<<foreach [tag in Tags]>>");
        builder.Writeln(" - <<[tag]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        // Save template.
        doc.Save(templatePath);

        // Load template for reporting.
        var reportDoc = new Document(templatePath);

        // Prepare data model.
        var model = new Model
        {
            Tags = new List<string> { "alpha", "beta", "gamma" }
        };

        // Build report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save final report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
