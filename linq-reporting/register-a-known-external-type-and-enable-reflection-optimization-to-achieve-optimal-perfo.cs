using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required for XML handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample XML data.
        const string xmlPath = "orders.xml";
        File.WriteAllText(xmlPath,
            @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Orders>
    <Order>
        <Id>1</Id>
        <CustomerName>John Doe</CustomerName>
        <OrderDate>2023-01-15</OrderDate>
    </Order>
    <Order>
        <Id>2</Id>
        <CustomerName>Jane Smith</CustomerName>
        <OrderDate>2023-02-20</OrderDate>
    </Order>
</Orders>");

        // Build a template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        // Call the static helper method to format the date.
        builder.Writeln("Date: <<[Utils.FormatDate(order.OrderDate)]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Create XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // Register the external type that will be used inside the template.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Utils));

        // Build the report. The data source name must match the root element used in the template.
        engine.BuildReport(doc, xmlDataSource, "Orders");

        // Save the generated report.
        doc.Save("report.docx");
    }
}

// External helper class with a static method that can be called from the template.
public static class Utils
{
    // Accepts either a string or a DateTime and formats it as yyyy-MM-dd.
    public static string FormatDate(object dateValue)
    {
        if (dateValue == null)
            return string.Empty;

        // If the value is already a DateTime, format it directly.
        if (dateValue is DateTime dt)
            return dt.ToString("yyyy-MM-dd");

        // Otherwise try to parse the string representation.
        if (DateTime.TryParse(dateValue.ToString(), out var parsed))
            return parsed.ToString("yyyy-MM-dd");

        // Fallback – return the original value as a string.
        return dateValue.ToString();
    }
}
