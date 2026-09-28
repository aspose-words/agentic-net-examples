using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple LINQ Reporting template.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title placeholder.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Table header.
        builder.Writeln("<<foreach [item in model.Items]>>");
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Table rows.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save template to a memory stream.
        using MemoryStream templateStream = new MemoryStream();
        templateDoc.Save(templateStream, SaveFormat.Docx);
        templateStream.Position = 0;

        // Prepare HTTP listener.
        HttpListener listener = new HttpListener();
        listener.Prefixes.Add("http://localhost:5000/");
        listener.Start();

        // Trigger a request so the listener does not block indefinitely.
        Task.Run(async () =>
        {
            using HttpClient client = new HttpClient();
            await client.GetStringAsync("http://localhost:5000/");
        });

        // Wait for a single request.
        HttpListenerContext context = listener.GetContext();

        // Build the report.
        using MemoryStream templateCopy = new MemoryStream(templateStream.ToArray());
        Document reportDoc = new Document(templateCopy);
        ReportingEngine engine = new ReportingEngine();

        ReportModel model = new ReportModel
        {
            Title = "Sample LINQ Reporting Output",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        engine.BuildReport(reportDoc, model, "model");

        // Save the generated PDF to a temporary memory stream first.
        using MemoryStream pdfStream = new MemoryStream();
        reportDoc.Save(pdfStream, SaveFormat.Pdf);
        pdfStream.Position = 0;

        // Write the PDF to the HTTP response stream.
        context.Response.ContentType = "application/pdf";
        context.Response.ContentLength64 = pdfStream.Length;
        pdfStream.CopyTo(context.Response.OutputStream);
        context.Response.OutputStream.Close();
        context.Response.Close();

        // Clean up.
        listener.Stop();
    }
}

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
