using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // 2. Load the template.
        Document doc = new Document(templatePath);

        // 3. Prepare the data model. The Name property will throw an exception.
        ReportModel model = new ReportModel
        {
            Product = new Product
            {
                // Price is valid, Name will throw.
                Price = 19.99m
            }
        };

        // 4. Configure the ReportingEngine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;
        engine.MissingMemberMessage = "[Missing]";

        // 5. Build the report inside a try‑catch block.
        try
        {
            // The engine will attempt to evaluate all expressions.
            // With InlineErrorMessages enabled, evaluation errors are inserted into the document
            // instead of propagating as exceptions.
            engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            // If an unexpected exception occurs, log it and continue.
            Console.WriteLine($"Report generation error: {ex.Message}");
        }

        // 6. Replace any inline error messages (or leftover tags) with a friendly placeholder.
        // The engine inserts messages that start with "Error evaluating expression".
        doc.Range.Replace(new Regex(@"Error evaluating expression.*"), "[Error]");
        // As a safety net, also replace any unreplaced LINQ tags.
        doc.Range.Replace(new Regex(@"<<\[[^\]]+\]>>"), "[Error]");

        // 7. Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("Name: <<[model.Product.Name]>>");
        builder.Writeln("Price: $<<[model.Product.Price]>>");

        doc.Save(filePath);
    }
}

// Wrapper class used as the root data source for the report.
public class ReportModel
{
    public Product Product { get; set; } = new Product();
}

// Sample data class. The Name getter throws to simulate a faulty expression.
public class Product
{
    public string Name
    {
        get
        {
            // Simulate an error during expression evaluation.
            throw new InvalidOperationException("Simulated property failure.");
        }
    }

    public decimal Price { get; set; }
}
