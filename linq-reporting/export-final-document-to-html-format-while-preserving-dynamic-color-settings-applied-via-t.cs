using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportModel
{
    public string Title { get; set; } = "Dynamic Title";
    public string ColorName { get; set; } = "Blue";
    public string Content { get; set; } = "This paragraph demonstrates a dynamic text color applied via LINQ Reporting tags.";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title with dynamic text color.
        builder.Writeln("<<textColor [model.ColorName]>><<[model.Title]>><</textColor>>");
        // Regular content.
        builder.Writeln("<<[model.Content]>>");

        // Save the template to disk (required by the workflow).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 3. Export the populated document to HTML, preserving colors.
        // -----------------------------------------------------------------
        var htmlOptions = new HtmlSaveOptions
        {
            // Ensure that the generated HTML keeps the original styling.
            ExportFontResources = true,
            ExportImagesAsBase64 = true,
            ExportTextInputFormFieldAsText = true
        };

        const string htmlPath = "Report.html";
        loadedTemplate.Save(htmlPath, htmlOptions);

        // The example finishes without waiting for user input.
    }
}
