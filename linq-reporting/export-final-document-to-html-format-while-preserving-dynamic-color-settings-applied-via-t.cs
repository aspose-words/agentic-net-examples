using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportModel
{
    // Sample title (not used directly in tags but can be added if needed)
    public string Title { get; set; } = "Dynamic Color Report";

    // Color name or HTML color code used by textColor and backColor tags
    public string ColorName { get; set; } = "Tomato";

    // HTML snippet that will be inserted using the -html switch
    public string HtmlSnippet { get; set; } = "<b>Bold HTML content</b> with <i>italic</i> text.";
}

public class Program
{
    public static void Main()
    {
        // Prepare the data model
        ReportModel model = new ReportModel();

        // Create a blank document that will serve as the LINQ Reporting template
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Paragraph with dynamic text color
        builder.Writeln("<<textColor [model.ColorName]>>This text is colored dynamically<</textColor>>");

        // Paragraph with dynamic background color
        builder.Writeln("<<backColor [model.ColorName]>>This paragraph has a dynamic background color<</backColor>>");

        // Insert an HTML fragment using the -html switch; the HTML will be preserved in the output
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(template, model, "model");

        // Prepare HTML save options to ensure colors are exported correctly
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            Encoding = Encoding.UTF8,
            ExportListLabels = ExportListLabels.Auto,
            ExportDocumentProperties = false,
            ExportImagesAsBase64 = false,
            ExportFontResources = false,
            ExportGeneratorName = true
        };

        // Ensure the output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the final document as HTML
        string htmlPath = Path.Combine(outputDir, "DynamicColorReport.html");
        template.Save(htmlPath, htmlOptions);
    }
}
