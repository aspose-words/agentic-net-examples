using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Total amount of the order.
    public double Total { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write the total amount placeholder.
        builder.Writeln("Order Total: <<[order.Total]>>");

        // Compute discount using conditional and arithmetic operators.
        // If Total > 500, discount = Total * 0.1, otherwise 0.
        builder.Writeln("Discount: <<[order.Total > 500 ? order.Total * 0.1 : 0]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for report generation.
        Document report = new Document(templatePath);

        // Step 3: Prepare the data model.
        Order order = new Order { Total = 750 }; // Example total exceeding 500.

        // Step 4: Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, order, "order");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}
