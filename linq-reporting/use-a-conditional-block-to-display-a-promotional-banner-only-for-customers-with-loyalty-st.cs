using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    // Customer name – initialized to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;

    // Indicates whether the customer has a loyalty status.
    public bool IsLoyal { get; set; }
}

public class ReportModel
{
    // Collection of customers that will be used as the data source.
    public List<Customer> Customers { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice Johnson",  IsLoyal = true  },
                new Customer { Name = "Bob Smith",      IsLoyal = false },
                new Customer { Name = "Carol Williams", IsLoyal = true  }
            }
        };

        // -----------------------------------------------------------------
        // 2. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";

        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the Customers collection.
        builder.Writeln("<<foreach [c in Customers]>>");

        // Output the customer's name.
        builder.Writeln("Customer: <<[c.Name]>>");

        // Conditional block – show a promotional banner only for loyal customers.
        builder.Writeln("<<if [c.IsLoyal]>>");
        // Optional styling: green text for the banner.
        builder.Writeln("<<textColor [\"Green\"]>>Loyalty Promotion!<</textColor>>");
        builder.Writeln("<</if>>");

        // End of the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        var engine = new ReportingEngine
        {
            // No special options are required for this simple example.
            Options = ReportBuildOptions.None
        };

        // Build the report using the model as the root data source.
        // The root name "model" must match the name used in the template tags.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Optional: check the success flag (relevant only when InlineErrorMessages is set).
        if (!success)
        {
            Console.WriteLine("Report generation encountered errors.");
        }

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Inform the user that the process completed.
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
