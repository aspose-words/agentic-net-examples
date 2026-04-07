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
            @"<?xml version=""1.0"" encoding=""utf-8""?>
              <Order>
                  <Amount>1234.56</Amount>
              </Order>");

        // Create a template document with a LINQ Reporting tag.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Amount: <<[order.Amount]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Set a custom number format provider to format currency fields.
        doc.FieldOptions.ResultFormatter = new CurrencyFormatter();

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "order");

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Custom formatter that formats numeric values as currency.
    private class CurrencyFormatter : IFieldResultFormatter
    {
        public string FormatNumeric(double value, string format)
        {
            // Apply currency formatting using the current culture.
            return value.ToString("C", CultureInfo.CurrentCulture);
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
