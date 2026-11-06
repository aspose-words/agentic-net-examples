using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model
        var model = new ReportModel
        {
            Title = "Demo Report",
            Tags = new List<string> { "Tag1", "Tag2", "Tag3" }
        };

        // Create template document
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[Title]>>");
        builder.Writeln("<<if [Tags.Any()]>>");
        builder.Writeln("Tags:");
        builder.Writeln("<<foreach [tag in Tags]>>");
        builder.Writeln("- <<[tag]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");
        templateDoc.Save(templatePath);

        // Load template and generate report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save final report
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

public class ReportModel
{
    public string Title { get; set; } = "";
    public List<string> Tags { get; set; } = new();
}
