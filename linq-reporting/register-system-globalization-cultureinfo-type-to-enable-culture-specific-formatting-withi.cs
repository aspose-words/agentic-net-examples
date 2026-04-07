using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // 1. Create a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert expressions that format a date and a number using specific cultures.
        // Note the correct LINQ Reporting tag syntax: <<[expression]>>.
        builder.Writeln(
            "Date (en-US): <<[order.Date.ToString(\"d\", CultureInfo.GetCultureInfo(\"en-US\"))]>>");
        builder.Writeln(
            "Date (fr-FR): <<[order.Date.ToString(\"d\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");
        builder.Writeln(
            "Number (en-US): <<[order.Amount.ToString(\"N\", CultureInfo.GetCultureInfo(\"en-US\"))]>>");
        builder.Writeln(
            "Number (fr-FR): <<[order.Amount.ToString(\"N\", CultureInfo.GetCultureInfo(\"fr-FR\"))]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare the data source.
        var order = new Order
        {
            Date = new DateTime(2023, 12, 31),
            Amount = 12345.67m
        };

        // 4. Configure the ReportingEngine.
        var engine = new ReportingEngine();

        // Register System.Globalization.CultureInfo so it can be used inside template expressions.
        engine.KnownTypes.Add(typeof(CultureInfo));

        // 5. Build the report. The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, order, "order");

        // 6. Save the generated report.
        doc.Save(reportPath);
    }
}

// Public data model used by the template.
public class Order
{
    public DateTime Date { get; set; } = DateTime.Now;
    public decimal Amount { get; set; } = 0m;
}
