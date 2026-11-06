using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        ReportModel model = new()
        {
            Title = "Sample Report"
        };

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new();
        DocumentBuilder builder = new(template);

        // Insert a LINQ Reporting tag that will be replaced with the title.
        builder.Writeln("<<[model.Title]>>");

        // Add a footer with the dynamic page number expression tag.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("{=PageNumber}");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(report, model, "model");

        // Update fields (e.g., the page number field) so they display correct values.
        report.UpdateFields();

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// Public data model class required by the LINQ Reporting engine.
public class ReportModel
{
    public string Title { get; set; } = "";
}
