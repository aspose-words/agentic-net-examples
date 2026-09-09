using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

public class ReportModel
{
    // The image stream may be null to simulate a missing image.
    public Stream? ImageStream { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and output documents.
        const string templatePath = "Template.docx";
        const string outputPath = "Output.docx";

        // -------------------------------------------------
        // 1. Create a template document containing an image tag.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Image tags must be placed inside a textbox.
        var textBox = builder.InsertShape(Aspose.Words.Drawing.ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);
        // The tag references the ImageStream property of the model.
        builder.Write("<<image [model.ImageStream]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back (required by the workflow).
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data model with a null image stream.
        // -------------------------------------------------
        var model = new ReportModel
        {
            ImageStream = null // Intentionally null to test graceful failure.
        };

        // -------------------------------------------------
        // 4. Build the report using InlineErrorMessages option.
        // -------------------------------------------------
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns false when an error occurs and InlineErrorMessages is set.
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // -------------------------------------------------
        // 5. Verify the result and output information.
        // -------------------------------------------------
        Console.WriteLine($"BuildReport succeeded: {success}");
        if (!success)
        {
            // The engine should have inserted an error message into the document.
            string documentText = loadedTemplate.GetText();
            Console.WriteLine("Document contains error message:");
            Console.WriteLine(documentText);
        }

        // Save the resulting document for inspection.
        loadedTemplate.Save(outputPath);
    }
}
