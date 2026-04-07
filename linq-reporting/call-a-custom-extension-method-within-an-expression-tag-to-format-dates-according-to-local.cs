using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DateExtensions
{
    // Extension method that formats a DateTime using the specified locale (culture name).
    // The ReportingEngine can invoke this method as a static method.
    public static string ToLocaleString(this DateTime date, string locale)
    {
        var culture = new CultureInfo(locale);
        return date.ToString(culture);
    }
}

public class Order
{
    public DateTime OrderDate { get; set; } = DateTime.Now;
    public string Customer { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -------------------- Create template document --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Orders collection.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Output customer name.
        builder.Writeln("Customer: <<[order.Customer]>>");

        // Use the static method call to format the date in French locale.
        // The ReportingEngine resolves static methods from types added to KnownTypes.
        builder.Writeln("Date (FR): <<[DateExtensions.ToLocaleString(order.OrderDate, \"fr-FR\")]>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, just for inspection).
        doc.Save("Template.docx");

        // -------------------- Prepare data source --------------------
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { Customer = "Alice", OrderDate = new DateTime(2023, 5, 10) },
                new Order { Customer = "Bob",   OrderDate = new DateTime(2023, 6, 15) }
            }
        };

        // -------------------- Build the report --------------------
        ReportingEngine engine = new ReportingEngine();

        // Register the static class that contains the extension method.
        engine.KnownTypes.Add(typeof(DateExtensions));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // -------------------- Save the final document --------------------
        doc.Save("Report.docx");
    }
}
