using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible legacy encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a simple data model.
        var product = new Product
        {
            Name = "Sample Widget",
            Price = 1234.56m
        };

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that reference the model.
        builder.Writeln("Product: <<[model.Name]>>");
        builder.Writeln("Price: <<[model.Price]>>"); // Number will be formatted using the current culture.

        // Configure a custom culture for number formatting.
        // Here we use French (France) which formats numbers with a space as thousands separator
        // and a comma as the decimal separator.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

        // Ensure that field number formatting also respects the current culture.
        template.FieldOptions.UseInvariantCultureNumberFormat = false;

        // Create and configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // No special options are required for this scenario.
            Options = ReportBuildOptions.None
        };

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(template, product, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Public data model class with properties referenced by the template.
public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}
