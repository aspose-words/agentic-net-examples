using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlPath = "order.xml";
        File.WriteAllText(xmlPath,
            @"<order>
                <Id>1001</Id>
                <Amount>1234.56</Amount>
                <Customer>John Doe</Customer>
              </order>");

        // Create a template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Apply a custom number format provider for currency fields.
        doc.FieldOptions.ResultFormatter = new CurrencyResultFormatter();

        // Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "order");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Custom formatter that formats numeric values as currency.
    private class CurrencyResultFormatter : IFieldResultFormatter
    {
        public string FormatNumeric(double value, string format)
        {
            // Apply custom currency format (e.g., $1,234.56).
            return string.Format(CultureInfo.InvariantCulture, "${0:N2}", value);
        }

        public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
        {
            // No custom date formatting required.
            return null;
        }

        public string Format(string value, GeneralFormat format)
        {
            // No custom general formatting required.
            return null;
        }

        public string Format(double value, GeneralFormat format)
        {
            // No custom general formatting required.
            return null;
        }
    }
}
