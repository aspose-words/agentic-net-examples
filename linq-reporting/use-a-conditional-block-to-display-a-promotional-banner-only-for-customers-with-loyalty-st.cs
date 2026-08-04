using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection "Customers".
        builder.Writeln("<<foreach [c in Customers]>>");
        // Output the customer's name.
        builder.Writeln("Customer: <<[c.Name]>>");
        // Conditional block: show a promotional banner only for loyal customers.
        builder.Writeln("<<if [c.IsLoyal]>>");
        // The banner text is highlighted in green.
        builder.Writeln("<<textColor [\"Green\"]>>Loyalty Promotion!<</textColor>>");
        builder.Writeln("<</if>>");
        // End of the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Build a realistic data source.
        ReportModel model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice", IsLoyal = true },
                new Customer { Name = "Bob", IsLoyal = false },
                new Customer { Name = "Charlie", IsLoyal = true }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Customer
{
    public string Name { get; set; } = string.Empty;
    public bool IsLoyal { get; set; }
}
