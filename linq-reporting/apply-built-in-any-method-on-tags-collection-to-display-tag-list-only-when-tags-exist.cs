using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<string> Tags { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // Step 1: Create the LINQ Reporting template.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Tag List:");

        // Show the list only when the collection has at least one element.
        builder.Writeln("<<if [model.Tags.Any()]>>");

        // Iterate over the tags.
        builder.Writeln("<<foreach [tag in model.Tags]>>- <<[tag]>>");
        builder.Writeln("<</foreach>>");

        // End of the conditional block.
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Step 2: Load the template and prepare data.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data with a few tags.
        ReportModel data = new ReportModel
        {
            Tags = new List<string> { "Alpha", "Beta", "Gamma" }
        };

        // -------------------------------------------------
        // Step 3: Build the report using the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, data, "model");

        // -------------------------------------------------
        // Step 4: Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
