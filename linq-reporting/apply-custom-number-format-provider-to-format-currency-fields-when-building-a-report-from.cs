using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

public class Program
{
    // Custom formatter that formats numeric values as currency.
    private class CurrencyFormatter : IFieldResultFormatter
    {
        private readonly string _currencyFormat;

        public CurrencyFormatter(string currencyFormat = "C")
        {
            _currencyFormat = currencyFormat;
        }

        // Format numeric values (e.g., fields like <<[order.Amount]>>).
        public string FormatNumeric(double value, string format)
        {
            // Use the provided currency format and invariant culture for consistency.
            return value.ToString(_currencyFormat, CultureInfo.InvariantCulture);
        }

        // The other methods are not needed for this example; return null to use default handling.
        public string FormatDateTime(DateTime value, string format, CalendarType calendarType) => null;
        public string Format(string value, GeneralFormat format) => null;
        public string Format(double value, GeneralFormat format) => null;
    }

    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // 1. Create a simple XML data source file.
        string xmlPath = Path.Combine(workDir, "Orders.xml");
        File.WriteAllText(xmlPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <Id>1001</Id>
        <CustomerName>John Doe</CustomerName>
        <Amount>1234.56</Amount>
    </Order>
    <Order>
        <Id>1002</Id>
        <CustomerName>Jane Smith</CustomerName>
        <Amount>7890.12</Amount>
    </Order>
</Orders>");

        // 2. Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header
        builder.Writeln("Order Report");
        builder.Writeln("------------------------------");

        // Repeat for each Order element.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        // The Amount field will be formatted by the custom formatter.
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Apply the custom number format provider.
        doc.FieldOptions.ResultFormatter = new CurrencyFormatter();

        // 5. Load the XML data source.
        var xmlDataSource = new Aspose.Words.Reporting.XmlDataSource(xmlPath);

        // 6. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, xmlDataSource, "Orders");

        // 7. Save the generated report.
        string outputPath = Path.Combine(workDir, "Report.docx");
        doc.Save(outputPath);
    }
}
