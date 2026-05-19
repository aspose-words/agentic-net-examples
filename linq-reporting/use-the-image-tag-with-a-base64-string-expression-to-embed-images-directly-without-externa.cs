using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Full data‑URI string (kept for reference)
    public string Base64Image { get; set; } =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ+XKcAAAAASUVORK5CYII=";

    // Returns only the raw Base64 part decoded to a byte array.
    // The LINQ Reporting engine accepts a byte[] for an image tag.
    public byte[] ImageBytes
    {
        get
        {
            // Split on the first comma to discard the "data:image/png;base64," prefix.
            var parts = Base64Image.Split(new[] { ',' }, 2);
            var base64 = parts.Length == 2 ? parts[1] : Base64Image;
            return Convert.FromBase64String(base64);
        }
    }
}

public class Program
{
    public static void Main()
    {
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a textbox that will host the image.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);

        // Image tag that reads the byte[] from the model.
        builder.Write("<<image [Model.ImageBytes] -fitSize>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Prepare the data model.
        // -------------------------------------------------
        var model = new ReportModel();

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object name in the template is "Model".
        engine.BuildReport(reportDoc, model, "Model");

        // Save the final document.
        reportDoc.Save(reportPath);
    }
}
