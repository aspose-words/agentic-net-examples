using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class Order
    {
        public decimal Price { get; set; } = 0m;
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Price: <<[order.Price]>>"); // Tag will be replaced by the model value.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure a custom culture for number formatting.
        //    The LINQ Reporting engine uses the current thread's culture
        //    when converting numeric values to strings.
        // -----------------------------------------------------------------
        CultureInfo customCulture = new CultureInfo("fr-FR"); // French uses comma as decimal separator.
        CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = customCulture;

        // -----------------------------------------------------------------
        // 4. Prepare the data source.
        // -----------------------------------------------------------------
        Order order = new Order { Price = 1234.56m };

        // -----------------------------------------------------------------
        // 5. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, order, "order");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        // Restore the original culture (optional cleanup).
        Thread.CurrentThread.CurrentCulture = originalCulture;
    }
}
