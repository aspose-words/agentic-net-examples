using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple data model with a numeric value.
        ReportModel model = new()
        {
            Price = 1234.56m
        };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new(templatePath);

        // Set a custom culture (German) for number formatting.
        CultureInfo customCulture = new("de-DE");
        Thread.CurrentThread.CurrentCulture = customCulture;

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Output the resulting text to the console (for demonstration purposes).
        Console.WriteLine("Report content:");
        Console.WriteLine(doc.GetText());
    }

    private static void CreateTemplate(string path)
    {
        Document template = new();
        DocumentBuilder builder = new(template);

        // Insert a LINQ Reporting tag that will display the price.
        // The numeric value will be formatted according to the current thread's culture.
        builder.Writeln("Price: <<[model.Price]>>");

        // Save the template.
        template.Save(path);
    }
}

public class ReportModel
{
    // Initialize with a non‑null default to avoid nullable warnings.
    public decimal Price { get; set; } = 0m;
}
