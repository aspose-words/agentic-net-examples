using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Data model that will be bound to the template.
    public class PriceReport
    {
        // Initialize the list to avoid nullable warnings.
        public List<decimal> Prices { get; set; } = new();

        // Expose the minimum price as a property for the template.
        public decimal MinPrice => Prices.Min();
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that displays the minimum price.
        // The tag uses the root object name "model" (will be supplied later).
        builder.Writeln("Discount benchmark price: <<[model.MinPrice]>>");

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data source.
        // -----------------------------------------------------------------
        PriceReport model = new PriceReport
        {
            Prices = new List<decimal> { 199.99m, 149.50m, 179.75m, 129.99m }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // BuildReport overload that allows referencing the root object name.
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
