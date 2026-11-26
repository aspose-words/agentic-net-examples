using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static async Task Main(string[] args)
    {
        // Ensure a sample image exists.
        const string imageFileName = "sample.png";
        if (!File.Exists(imageFileName))
        {
            // 1x1 transparent PNG.
            const string pngBase64 =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6UAAAAASUVORK5CYII=";
            var imageBytes = Convert.FromBase64String(pngBase64);
            await File.WriteAllBytesAsync(imageFileName, imageBytes);
        }

        // Start a simple HTTP listener.
        var listener = new HttpListener();
        listener.Prefixes.Add("http://localhost:5000/report/");
        listener.Start();

        // Handle a single request.
        var handlingTask = Task.Run(async () =>
        {
            var context = await listener.GetContextAsync();
            var response = context.Response;
            response.ContentType =
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

            // Build the model.
            var model = new ReportModel
            {
                Title = "LINQ Reporting Sample",
                HtmlSnippet = "<b>Bold HTML Content</b>",
                ImagePath = Path.GetFullPath(imageFileName),
                BookmarkName = "SampleBookmark",
                Url = "https://example.com",
                LinkText = "Example Site",
                ColorName = "Blue"
            };

            // Create the template document.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<<[model.HtmlSnippet] -html>>");
            builder.Writeln("<<bookmark [model.BookmarkName]>>");
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<<link [model.Url] [model.LinkText]>>");
            builder.Writeln("<<textColor [model.ColorName]>>Colored Text<</textColor>>");

            // Image inside a textbox.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 120);
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("<<image [model.ImagePath] -fitSize>>");

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save to a memory stream first.
            using var ms = new MemoryStream();
            templateDoc.Save(ms, SaveFormat.Docx);
            var resultBytes = ms.ToArray();

            // Write the result to the HTTP response.
            response.ContentLength64 = resultBytes.Length;
            await response.OutputStream.WriteAsync(resultBytes, 0, resultBytes.Length);
            response.OutputStream.Close();
            response.Close();
        });

        // Small pause to ensure the listener is ready.
        await Task.Delay(200);

        // Request the report.
        using var httpClient = new HttpClient();
        var reportBytes = await httpClient.GetByteArrayAsync("http://localhost:5000/report/");
        await File.WriteAllBytesAsync("report_output.docx", reportBytes);

        // Wait for the handling task to finish and then stop the listener.
        await handlingTask;
        listener.Stop();
    }
}

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string HtmlSnippet { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
    public string ColorName { get; set; } = string.Empty;
}
