using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Extensions
{
    // Formats a decimal value as a currency string with a dollar sign and two decimal places.
    public static string ToCurrencyString(this decimal amount)
    {
        return string.Format(CultureInfo.InvariantCulture, "${0:0.00}", amount);
    }

    // Helper method for the reporting engine (static call) – required because the engine
    // cannot resolve extension methods directly in template expressions.
    public static string ToCurrencyStringStatic(decimal amount)
    {
        return ToCurrencyString(amount);
    }
}

// Simple data model that will be used as the root object for the report.
public class Order
{
    public decimal Amount { get; set; } = 0m; // Initialized to avoid nullable warnings.
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to insert the template tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that calls the static helper method.
        // The engine can resolve static methods from known types, so we use the static wrapper.
        builder.Writeln("Amount: <<[Extensions.ToCurrencyStringStatic(order.Amount)]>>");

        // Prepare sample data.
        Order order = new Order { Amount = 1234.56m };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the static class that contains the helper method so the engine can use it.
        engine.KnownTypes.Add(typeof(Extensions));

        // Build the report. The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, order, "order");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
