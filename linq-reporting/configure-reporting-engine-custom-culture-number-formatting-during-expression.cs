using System;
using System.Globalization;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template document that contains reporting expressions.
        Document doc = new Document("Template.docx");

        // Set a custom culture (for example, French) that will be used for number formatting.
        Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");

        // Ensure that fields use the current thread's culture rather than the invariant culture.
        doc.FieldOptions.UseInvariantCultureNumberFormat = false;

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Example data source for the report.
        var data = new { Amount = 12345.67 };

        // Build the report using the template and the data source.
        engine.BuildReport(doc, data);

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
