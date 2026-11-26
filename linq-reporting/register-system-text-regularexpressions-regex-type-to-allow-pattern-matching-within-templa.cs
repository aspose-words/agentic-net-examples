using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Email { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // Create template document with LINQ Reporting tags
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Email: <<[model.Email]>>");
        builder.Writeln(
            "<<if [Regex.IsMatch(model.Email, \"^\\\\S+@\\\\S+\\\\.\\\\S+$\")]>>" +
            "<<textColor [\"Red\"]>>Invalid Email<</textColor>><</if>>");

        template.Save(templatePath);

        // Prepare data model
        ReportModel model = new ReportModel
        {
            Email = "invalid-email" // Change to a valid email to see no warning
        };

        // Load template and build report
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
        Console.WriteLine($"BuildReport success: {success}");
    }
}
