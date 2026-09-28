using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public List<string> Tags { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create the template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Tag List:");
        builder.Writeln("<<if [Tags.Any()]>>");
        builder.Writeln("<<foreach [tag in Tags]>>");
        builder.Writeln("- <<[tag]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new Model
        {
            Tags = new() { "alpha", "beta", "gamma" }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
