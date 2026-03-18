using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the template document that contains the reporting tags.
        Document doc = new Document("Template.docx");

        // Attach a custom result formatter that will format all numeric field results as currency.
        doc.FieldOptions.ResultFormatter = new CurrencyResultFormatter();

        // Create an XML data source from the file that holds the report data.
        XmlDataSource xmlSource = new XmlDataSource("Data.xml");

        // Build the report by merging the XML data into the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlSource, "ds");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    // Implements IFieldResultFormatter to provide custom formatting for numeric fields.
    private class CurrencyResultFormatter : IFieldResultFormatter
    {
        // Called for numeric fields (\\# switch). Formats the value as currency.
        public string FormatNumeric(double value, string format)
        {
            // Use invariant culture to ensure consistent currency formatting.
            // Example format: $1,234.56
            return string.Format(CultureInfo.InvariantCulture, "${0:#,##0.00}", value);
        }

        // No custom date/time formatting required; return null to use default behavior.
        public string FormatDateTime(DateTime value, string format, CalendarType calendarType) => null;

        // No custom general formatting required; return null to use default behavior.
        public string Format(string value, GeneralFormat format) => null;
        public string Format(double value, GeneralFormat format) => null;
    }
}
