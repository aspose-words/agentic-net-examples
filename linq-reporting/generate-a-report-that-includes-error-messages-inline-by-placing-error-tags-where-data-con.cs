using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var builder = new DocumentBuilder();

        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // The -format switch forces a conversion; the value returned below cannot be formatted as a number,
        // so the ReportingEngine will insert an <<error>> tag when InlineErrorMessages is enabled.
        builder.Writeln("Faulty Value: <<[order.FaultyValue] -format \"N2\">>");

        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var order = new Order
        {
            CustomerName = "John Doe"
            // FaultyValue is intentionally invalid for numeric formatting.
        };

        // -----------------------------------------------------------------
        // 4. Build the report with inline error messages.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        bool success = engine.BuildReport(doc, order, "order");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "report.docx";
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 6. Output the result status.
        // -----------------------------------------------------------------
        Console.WriteLine($"Report generation success: {success}");
    }
}

// ---------------------------------------------------------------------
// Data model.
// ---------------------------------------------------------------------
public class Order
{
    // Non‑nullable property initialized to avoid warnings.
    public string CustomerName { get; set; } = string.Empty;

    // This property returns a value that cannot be formatted as a number,
    // causing a conversion error that will be shown inline as <<error>>.
    public string FaultyValue => "NotANumber";
}
