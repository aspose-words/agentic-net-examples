using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure code page support (required by Aspose.Words for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data file
        string xmlPath = "orders.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
  <Order>
    <CustomerName>John Doe</CustomerName>
    <Total>1234.56</Total>
  </Order>
  <Order>
    <CustomerName>Jane Smith</CustomerName>
    <Total>789.00</Total>
  </Order>
</Orders>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Load XML and map to model objects
        XDocument xDoc = XDocument.Load(xmlPath);
        var model = new ReportModel
        {
            Orders = new List<Order>()
        };
        foreach (var elem in xDoc.Root!.Elements("Order"))
        {
            var order = new Order
            {
                CustomerName = (string)elem.Element("CustomerName")!,
                Total = decimal.Parse((string)elem.Element("Total")!)
            };
            model.Orders.Add(order);
        }

        // 3. Create a Word template with LINQ Reporting tags
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.FormattedTotal]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 4. Load the template and build the report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report
        string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Output file locations (no interactive prompts)
        Console.WriteLine($"Template saved to: {Path.GetFullPath(templatePath)}");
        Console.WriteLine($"Report generated at: {Path.GetFullPath(reportPath)}");
    }
}

// Custom format provider for currency formatting
public class CustomCurrencyFormatProvider : IFormatProvider, ICustomFormatter
{
    private readonly string _currencySymbol;

    public CustomCurrencyFormatProvider(string currencySymbol = "¤")
    {
        _currencySymbol = currencySymbol;
    }

    public object GetFormat(Type? formatType)
    {
        return formatType == typeof(ICustomFormatter) ? this : null!;
    }

    public string Format(string? format, object? arg, IFormatProvider? formatProvider)
    {
        if (string.IsNullOrEmpty(format) || arg == null)
            return string.Empty;

        // Handle custom currency format "C"
        if (format.Equals("C", StringComparison.OrdinalIgnoreCase) && arg is IFormattable formattable)
        {
            // Use invariant culture for number formatting, then prepend custom symbol
            return $"{_currencySymbol} {formattable.ToString("N2", CultureInfo.InvariantCulture)}";
        }

        // Fallback to default formatting
        if (arg is IFormattable defaultFormattable)
            return defaultFormattable.ToString(format, CultureInfo.CurrentCulture);

        return arg.ToString() ?? string.Empty;
    }
}

// Data model classes
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Total { get; set; }

    // Formatted total using the custom currency format provider
    public string FormattedTotal => Total.ToString("C", new CustomCurrencyFormatProvider("USD"));
}
