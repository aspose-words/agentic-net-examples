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
        // Register code page provider for Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create sample XML data file.
        const string xmlPath = "Products.xml";
        File.WriteAllText(xmlPath,
@"<Products>
    <Product>
        <Name>Apple</Name>
        <Price>1.23</Price>
    </Product>
    <Product>
        <Name>Banana</Name>
        <Price>0.75</Price>
    </Product>
    <Product>
        <Name>Cherry</Name>
        <Price>2.50</Price>
    </Product>
</Products>");

        // Load XML and populate the model.
        XDocument doc = XDocument.Load(xmlPath);
        var model = new ReportModel
        {
            Products = new List<Product>(),
            CurrencyProvider = new MyCurrencyFormatProvider()
        };

        foreach (var elem in doc.Root?.Elements("Product") ?? new List<XElement>())
        {
            model.Products.Add(new Product
            {
                Name = (string?)elem.Element("Name") ?? string.Empty,
                Price = decimal.TryParse((string?)elem.Element("Price"), out var p) ? p : 0m
            });
        }

        // Create a Word document template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Begin foreach loop over products.
        builder.Writeln("<<foreach [p in model.Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Price: <<[p.Price.ToString(\"C\", model.CurrencyProvider)]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the result.
        const string outputPath = "Report.docx";
        templateDoc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Model classes.
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
    public IFormatProvider CurrencyProvider { get; set; } = new MyCurrencyFormatProvider();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
}

// Custom currency format provider.
public class MyCurrencyFormatProvider : IFormatProvider
{
    private readonly CultureInfo _culture;

    public MyCurrencyFormatProvider()
    {
        _culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        _culture.NumberFormat.CurrencySymbol = "$";
        _culture.NumberFormat.CurrencyDecimalDigits = 2;
    }

    public object GetFormat(Type? formatType)
    {
        return _culture.GetFormat(formatType!);
    }
}
