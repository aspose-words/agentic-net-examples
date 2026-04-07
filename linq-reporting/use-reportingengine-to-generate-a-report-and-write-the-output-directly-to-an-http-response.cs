using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportItem
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<ReportItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Simple title tag.
        builder.Writeln("Report Title: <<[model.Title]>>");

        // Loop over the collection of items.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // 2. Prepare sample data that matches the tags used in the template.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Title = "Sample Report",
            Items = new List<ReportItem>
            {
                new ReportItem { Index = 1, Name = "Alpha" },
                new ReportItem { Index = 2, Name = "Beta" },
                new ReportItem { Index = 3, Name = "Gamma" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The third parameter ("model") makes the root object accessible in the template.
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated document directly to a stream.
        //    In a real web application this stream would be the HTTP response
        //    output stream. Here we use a MemoryStream to simulate that.
        // -----------------------------------------------------------------
        using (MemoryStream outputStream = new MemoryStream())
        {
            // Save the document in DOCX format to the stream.
            template.Save(outputStream, SaveFormat.Docx);

            // For demonstration, show the size of the generated document.
            Console.WriteLine($"Document written to stream. Length: {outputStream.Length} bytes.");
        }
    }
}
