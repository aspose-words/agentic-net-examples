using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Prepare temporary files.
        string templatePath = Path.Combine(Path.GetTempPath(), "Template.docx");
        string dataPath = Path.Combine(Path.GetTempPath(), "Data.xml");
        string outputPath = Path.Combine(Path.GetTempPath(), "Report.docx");

        // Create a simple template document with a numeric merge field.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Price:");
        // The field uses the reporting tag syntax <<ds.Price>> with a numeric format switch.
        builder.InsertField("MERGEFIELD ds.Price \\# \"0.00\"", null);
        doc.Save(templatePath);

        // Create a minimal XML data source.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<ds>
    <Price>1234.56</Price>
</ds>";
        File.WriteAllText(dataPath, xmlContent);

        // Load the template document.
        Document reportDoc = new Document(templatePath);

        // Attach a custom result formatter that will format all numeric field results as currency.
        reportDoc.FieldOptions.ResultFormatter = new CurrencyResultFormatter();

        // Create an XML data source from the file that holds the report data.
        XmlDataSource xmlSource = new XmlDataSource(dataPath);

        // Build the report by merging the XML data into the template.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlSource, "ds");

        // Save the generated report.
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
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
