using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    // Sample properties – initialize to avoid nullable warnings
    public string Name { get; set; } = "John Doe";
    public bool IsLoyal { get; set; } = true;
    public string LoyaltyLevel { get; set; } = "Gold";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with a conditional block.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Greeting
        builder.Writeln("Dear <<[customer.Name]>>,");
        builder.Writeln();

        // Conditional promotional banner – shown only when IsLoyal is true
        builder.Writeln("<<if [customer.IsLoyal]>>");
        builder.Writeln("=== Exclusive Offer for <<[customer.LoyaltyLevel]>> Members! ===");
        builder.Writeln("Get 20% off on your next purchase.");
        builder.Writeln("<</if>>");

        builder.Writeln();
        builder.Writeln("Thank you for being with us.");

        // Save the template to disk (required by the workflow)
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (simulating a real‑world scenario)
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data
        // -----------------------------------------------------------------
        Customer customer = new Customer
        {
            Name = "Alice Smith",
            IsLoyal = true,               // Change to false to hide the banner
            LoyaltyLevel = "Platinum"
        };

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are needed for this simple example
        engine.BuildReport(doc, customer, "customer");

        // -----------------------------------------------------------------
        // 5. Save the generated report
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
