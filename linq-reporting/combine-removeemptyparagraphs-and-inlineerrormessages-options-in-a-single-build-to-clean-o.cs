using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    // Sample data – non‑nullable properties are initialized to avoid warnings.
    public string CustomerName { get; set; } = "John Doe";

    // This property is intentionally left out of the class to trigger an inline error message.
    // The template will reference <<[order.MissingProp]>> which does not exist.

    // Empty string – after the tag is processed the paragraph will be empty.
    public string EmptyProp { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Paragraph that will be filled with a real value.
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // Paragraph that references a missing member – will produce an inline error message.
        builder.Writeln("Missing: <<[order.MissingProp]>>");

        // Paragraph that resolves to an empty string – will become empty and be removed.
        builder.Writeln("Empty: <<[order.EmptyProp]>>");

        // Save the template to disk (required before BuildReport according to the rules).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back and build the report.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Sample data source.
        Order order = new Order();

        // Configure the reporting engine to remove empty paragraphs and inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages;

        // Build the report. The root object name must match the tag prefix ("order").
        bool success = engine.BuildReport(loadedTemplate, order, "order");

        // Save the generated report.
        loadedTemplate.Save(reportPath);

        // Output the result flag – true means the template was parsed successfully (errors are inlined).
        Console.WriteLine($"Report built successfully: {success}");
    }
}
