using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD that formats a date using the \\@ switch.
        // The field will later be populated by the reporting engine.
        builder.InsertField(@"MERGEFIELD Date \@ ""dddd, d MMMM yyyy""");

        // -----------------------------------------------------------------
        // Register System.Globalization.CultureInfo so that it can be used
        // inside template expressions (e.g., CultureInfo.GetCultureInfo("fr-FR")).
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(CultureInfo));

        // Prepare a simple data source containing a DateTime value.
        var data = new { Date = new DateTime(2023, 5, 15) };

        // Build the report – the engine will replace the MERGEFIELD with the
        // value from the data source.
        engine.BuildReport(doc, data);

        // Instruct the document to obtain the culture for field formatting
        // from the field code (the \\@ switch). This allows the field to use
        // the culture that may be specified in the template expression.
        doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;

        // Update all fields so that the date is formatted according to the
        // culture information available at runtime.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Report.docx");
    }
}
