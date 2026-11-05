using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data (not used directly by the reporting engine in this example,
        // but kept to illustrate the original source).
        string xmlPath = "orders.xml";
        File.WriteAllText(xmlPath,
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Total>1234.56</Total>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Total>7890.12</Total>
    </Order>
</Orders>");

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("==============");
        builder.Writeln();
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.TotalFormatted]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Create a custom number format provider (custom currency symbol).
        NumberFormatInfo customNfi = (NumberFormatInfo)CultureInfo.InvariantCulture.NumberFormat.Clone();
        customNfi.CurrencySymbol = "€";
        customNfi.CurrencyDecimalDigits = 2;
        CultureInfo customCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        customCulture.NumberFormat = customNfi;
        CultureInfo.CurrentCulture = customCulture;

        // Build the data model.
        ReportModel model = new ReportModel
        {
            Orders = new List<Order>
            {
                new Order { CustomerName = "John Doe", Total = 1234.56 },
                new Order { CustomerName = "Jane Smith", Total = 7890.12 }
            }
        };

        // Build the report using the data model.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public double Total { get; set; }

    // Returns the total formatted as currency using the current culture.
    public string TotalFormatted => Total.ToString("C", CultureInfo.CurrentCulture);
}
