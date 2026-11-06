using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public DateTime OrderDate { get; set; } = DateTime.Now;
}

// Helper class exposing the date‑formatting method as an instance member.
public class ExtensionsHelper
{
    public string FormatDate(DateTime date, string locale)
    {
        var culture = new CultureInfo(locale);
        return date.ToString(culture);
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with a LINQ Reporting expression.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // The expression calls the helper instance (named "ext") to format the date.
        builder.Writeln("Order date (en-US): <<[ext.FormatDate(order.OrderDate, \"en-US\")]>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var template = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        var order = new Order
        {
            OrderDate = new DateTime(2023, 12, 31, 15, 45, 0)
        };

        var ext = new ExtensionsHelper();

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();

        // Pass both the data source ("order") and the helper ("ext").
        engine.BuildReport(template, new object[] { order, ext }, new[] { "order", "ext" });

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        template.Save(outputPath);
    }
}
