using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Collection of tags to be displayed.
    public List<string> Tags { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // If the Tags collection has any items, display a heading and list them.
        builder.Writeln("<<if [model.Tags.Any()]>>Tags:");
        builder.Writeln("<<foreach [tag in model.Tags]>> - <<[tag]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Tags = new List<string> { "Alpha", "Beta", "Gamma" }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model; the root name is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}
