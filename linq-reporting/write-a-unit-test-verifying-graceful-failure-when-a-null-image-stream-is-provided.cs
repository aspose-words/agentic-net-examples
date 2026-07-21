using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;

#nullable enable

public class ImageStreamNullTest
{
    // Model class used by the LINQ Reporting template.
    public class ReportModel
    {
        // The image stream that will be referenced by the <<image>> tag.
        // For this test we intentionally set it to null to verify graceful failure.
        public Stream? ImageStream { get; set; } = null;
    }

    public static void Main()
    {
        // Create a blank Word document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox to host the image tag (required by the image tag rules).
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag that references the ImageStream property.
        // The tag must be placed inside the textbox.
        builder.Write("<<image [model.ImageStream]>>");

        // Prepare the data model with a null image stream.
        ReportModel model = new ReportModel();

        // Configure the ReportingEngine to output inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns false when parsing fails and InlineErrorMessages is enabled.
        bool success = engine.BuildReport(doc, model, "model");

        // Output the result. Expecting false because the image stream is null.
        Console.WriteLine($"Report build success: {success}");

        // Save the resulting document so you can inspect the inline error message if needed.
        doc.Save("ImageStreamNullTestResult.docx");
    }
}
