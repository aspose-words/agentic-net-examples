using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Create a template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("  <<[item.Index]>> - <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, order);

        // Set up a simple HTTP listener.
        var listener = new HttpListener();
        listener.Prefixes.Add("http://localhost:8080/");
        listener.Start();

        // Task that handles the incoming HTTP request and writes the document to the response stream.
        var responseTask = Task.Run(() =>
        {
            var context = listener.GetContext(); // blocks until a request arrives
            var response = context.Response;

            response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            response.AddHeader("Content-Disposition", "attachment; filename=Report.docx");

            // Save the document to a seekable memory stream first, then copy to the non‑seekable response stream.
            using (var ms = new MemoryStream())
            {
                doc.Save(ms, SaveFormat.Docx);
                ms.Position = 0; // reset position before reading
                ms.CopyTo(response.OutputStream);
            }

            response.OutputStream.Close();
            response.Close();
        });

        // Send a request to the listener so that it does not wait indefinitely.
        var requestTask = Task.Run(async () =>
        {
            using var client = new HttpClient();
            var resp = await client.GetAsync("http://localhost:8080/");
            await resp.Content.ReadAsByteArrayAsync(); // ensure the response is fully read
        });

        // Wait for both tasks to finish.
        Task.WaitAll(responseTask, requestTask);
        listener.Stop();
    }
}

// Data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
