using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple title placeholder.
        builder.Writeln("<<[model.Title]>>");

        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // Image tag that expects a Stream from the model.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back before building the report.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new()
        {
            Title = "Custom Image Loader Example"
        };

        // -------------------------------------------------
        // 4. Build the report using LINQ Reporting Engine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the template. The ImageStream property returns
// a small PNG image from an in‑memory byte array, avoiding external
// network calls that can fail at runtime.
// ---------------------------------------------------------------------
public class ReportModel
{
    public string Title { get; set; } = string.Empty;

    // Returns a Stream containing a 1×1 pixel PNG image.
    public Stream ImageStream
    {
        get
        {
            // Base64‑encoded PNG (1×1 transparent pixel).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            // Aspose.Words will close the stream after use.
            return new MemoryStream(imageBytes);
        }
    }
}
