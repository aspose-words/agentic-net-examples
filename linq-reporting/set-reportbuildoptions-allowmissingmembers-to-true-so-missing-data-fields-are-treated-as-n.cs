using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a missing field tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Missing field: <<[order.MissingField]>>");

        // Save and reload the template to simulate typical workflow.
        const string templatePath = "Template.docx";
        doc.Save(templatePath);
        Document template = new Document(templatePath);

        // Prepare data model that does NOT contain the MissingField property.
        Order order = new Order { CustomerName = "John Doe" };

        // Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report.
        engine.BuildReport(template, order, "order");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Data model used as the root object for the report.
public class Order
{
    public string CustomerName { get; set; } = "";
    // Note: No MissingField property is defined.
}
