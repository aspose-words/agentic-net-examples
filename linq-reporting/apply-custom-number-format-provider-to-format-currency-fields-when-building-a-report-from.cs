using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Create sample XML data.
        const string xmlFileName = "orders.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Total>1234.56</Total>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Total>7890.12</Total>
    </Order>
</Orders>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Load XML and map to objects.
        var xDoc = XDocument.Load(xmlFileName);
        var model = new ReportModel
        {
            Orders = new List<Order>()
        };
        foreach (var elem in xDoc.Root!.Elements("Order"))
        {
            var order = new Order
            {
                CustomerName = (string)elem.Element("CustomerName")!,
                Total = (decimal)elem.Element("Total")!
            };
            model.Orders.Add(order);
        }

        // Create template document programmatically.
        const string templateFileName = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.FormattedTotal]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templateFileName);

        // Load the template and build the report.
        var template = new Document(templateFileName);
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        const string outputFileName = "report.docx";
        template.Save(outputFileName);
    }
}

// Wrapper model for the report.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Data model representing an order.
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Total { get; set; }

    // Returns the total formatted with a custom currency format provider.
    public string FormattedTotal => Total.ToString("C", new MyCurrencyFormatProvider());
}

// Custom format provider that uses a custom currency symbol.
public class MyCurrencyFormatProvider : IFormatProvider, ICustomFormatter
{
    public object? GetFormat(Type? formatType) =>
        formatType == typeof(ICustomFormatter) ? this : null;

    public string Format(string? format, object? arg, IFormatProvider? provider)
    {
        if (arg is null)
            return string.Empty;

        // Use custom formatting only for currency ("C") or when format is null/empty.
        if (string.IsNullOrEmpty(format) || format.Equals("C", StringComparison.OrdinalIgnoreCase))
        {
            const string customSymbol = "¤";

            if (arg is IFormattable formattable)
            {
                // Format number with two decimal places and invariant culture.
                var number = formattable.ToString("N2", CultureInfo.InvariantCulture);
                return $"{customSymbol}{number}";
            }

            return $"{customSymbol}{arg}";
        }

        // Fallback to default formatting.
        if (arg is IFormattable defaultFormattable)
            return defaultFormattable.ToString(format, provider);

        return arg.ToString() ?? string.Empty;
    }
}
