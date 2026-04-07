using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple template with a textbox that contains an image tag.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template for reporting.
        Document template = new Document(templatePath);

        // Prepare the data model. The ImageStream property will fetch the image from a web service.
        ReportModel model = new ReportModel
        {
            // Use a publicly accessible image URL.
            ImageUrl = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png"
        };

        // Ensure the stream is positioned at the beginning before the engine reads it.
        model.ImageStream.Position = 0;

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("ReportOutput.docx");
    }

    // Creates a Word document that contains a textbox with an image tag.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a textbox shape to host the image tag.
        Shape shape = builder.InsertShape(ShapeType.TextBox, 300, 200);

        // Move the cursor inside the textbox.
        builder.MoveTo(shape.FirstParagraph);

        // Write the LINQ Reporting image tag. The expression refers to the ImageStream property of the root object.
        builder.Writeln("<<image [ImageStream]>>");

        doc.Save(filePath);
    }
}

// Data model used as the root object for the report.
// The ImageStream property returns a MemoryStream containing the image bytes fetched from a URL.
public class ReportModel
{
    // URL of the image to load.
    public string ImageUrl { get; set; } = string.Empty;

    private MemoryStream? _imageStream;

    // The stream that Aspose.Words will use for the <<image>> tag.
    public Stream ImageStream
    {
        get
        {
            if (_imageStream == null)
            {
                // Synchronously download the image data.
                using HttpClient client = new HttpClient();
                byte[] data = client.GetByteArrayAsync(ImageUrl).GetAwaiter().GetResult();
                _imageStream = new MemoryStream(data);
            }

            // Return a fresh stream positioned at the start.
            _imageStream.Position = 0;
            return _imageStream;
        }
    }
}
