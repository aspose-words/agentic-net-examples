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
        // Register code page provider for additional encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data source.
        const string xmlPath = "orders.xml";
        File.WriteAllText(xmlPath,
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Amount>1234.56</Amount>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Amount>7890.12</Amount>
    </Order>
</Orders>");

        // 2. Create a template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Orders Report");
        builder.Writeln("----------------------");
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        var doc = new Document(templatePath);

        // 4. Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // 5. Build the report using ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "orders");

        // 6. Apply custom number format provider for currency fields.
        doc.FieldOptions.ResultFormatter = new CurrencyResultFormatter();

        // 7. Update fields to apply the custom formatting.
        doc.UpdateFields();

        // 8. Save the final report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Custom formatter that formats numeric values as currency.
public class CurrencyResultFormatter : IFieldResultFormatter
{
    public string FormatNumeric(double value, string format)
    {
        // Format all numeric values as currency with two decimal places.
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
