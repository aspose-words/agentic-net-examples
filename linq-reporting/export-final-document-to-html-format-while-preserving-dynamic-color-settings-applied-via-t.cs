using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for HTML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple data model with dynamic color and text.
        var model = new ReportModel
        {
            ColorName = "Blue",
            Text = "This text is rendered with a dynamic color."
        };

        // -----------------------------------------------------------------
        // Step 1: Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a textColor tag that uses the model's ColorName property.
        // The inner <<[model.Text]>> tag inserts the actual text.
        builder.Writeln("<<textColor [model.ColorName]>><<[model.Text]>><</textColor>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // Step 3: Export the final document to HTML, preserving colors.
        // -----------------------------------------------------------------
        const string htmlOutputPath = "output.html";
        loadedTemplate.Save(htmlOutputPath, SaveFormat.Html);
    }
}

// Public data model class required by the reporting engine.
public class ReportModel
{
    // Color name can be any known color string (e.g., "Red", "Green", "#FF00FF").
    public string ColorName { get; set; } = "Black";

    // Text to display inside the colored span.
    public string Text { get; set; } = string.Empty;
}
