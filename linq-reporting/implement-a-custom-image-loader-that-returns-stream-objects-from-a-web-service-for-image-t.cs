using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class ImageLoader
{
    // Downloads an image from the given URL and returns a fresh MemoryStream.
    public static Stream LoadImage(string url)
    {
        using var httpClient = new HttpClient();
        byte[] data = httpClient.GetByteArrayAsync(url).Result;
        var stream = new MemoryStream(data);
        stream.Position = 0;
        return stream;
    }
}

public class ReportModel
{
    // URL of the image to be loaded from a web service.
    public string ImageUrl { get; set; } = "https://via.placeholder.com/150";

    // Returns a new Stream each time it is accessed.
    public Stream ImageStream => ImageLoader.LoadImage(ImageUrl);
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a textbox shape to host the image tag.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
        builder.MoveTo(textBox.FirstParagraph);

        // Write the LINQ Reporting image tag that uses the Stream property.
        builder.Write("<<image [model.ImageStream] -fitSize>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (simulating a real-world scenario).
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
