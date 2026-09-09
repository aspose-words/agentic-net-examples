using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // 1. Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report with image loaded from a custom image loader:");
        // Insert a textbox that will host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 200);
        builder.MoveTo(textBox.FirstParagraph);
        // The image tag expects a Stream returned by the expression.
        builder.Write("<<image [model.Image] -fitSize>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // 2. Load the saved template.
        Document reportDoc = new Document(templatePath);

        // 3. Prepare the data model. The Image property returns a fresh Stream
        //    containing the image bytes supplied from a local source.
        ReportModel model = new();

        // 4. Build the report using Aspose.Words LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// Data model used by the template. The Image property supplies a Stream.
public class ReportModel
{
    // Returns a new MemoryStream each time it is accessed.
    public Stream Image
    {
        get
        {
            // A 1x1 pixel PNG (transparent) encoded in base64.
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YcK5VIAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            return new MemoryStream(imageBytes);
        }
    }
}
