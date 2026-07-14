using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words on .NET Core)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data
        const string xmlContent = @"
<Orders>
    <Order Id='1' Customer='Alice'>
        <Service Name='Consulting' />
        <Service Name='Support' />
    </Order>
    <Order Id='2' Customer='Bob'>
        <Service Name='Development' />
    </Order>
    <Order Id='3' Customer='Anna'>
        <Service Name='Design' />
        <Service Name='Testing' />
        <Service Name='Deployment' />
    </Order>
</Orders>";
        const string xmlPath = "Orders.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Load XML and filter orders (customers whose name starts with 'A')
        XDocument xDoc = XDocument.Load(xmlPath);
        List<Order> filteredOrders = xDoc.Root!
            .Elements("Order")
            .Where(o => ((string?)o.Attribute("Customer"))?.StartsWith("A") == true)
            .Select(o => new Order
            {
                Id = (int?)o.Attribute("Id") ?? 0,
                Customer = (string?)o.Attribute("Customer") ?? string.Empty,
                Services = o.Elements("Service")
                            .Select(s => new Service { Name = (string?)s.Attribute("Name") ?? string.Empty })
                            .ToList()
            })
            .ToList();

        // 3. Prepare the root model for the report
        var model = new ReportModel { Orders = filteredOrders };

        // 4. Create the LINQ Reporting template programmatically
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln(); // empty line

        // Begin foreach over orders
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");

        // Numbered list of services for each order
        builder.ListFormat.ApplyNumberDefault(); // apply numbering to the next paragraph
        builder.Writeln("<<restartNum>><<foreach [svc in order.Services]>> <<[svc.Name]>> <</foreach>>");

        // End foreach over orders
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // 5. Build the report using the LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // 6. Save the generated report
        const string reportPath = "Report.docx";
        templateDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}

// Root wrapper class
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Order class
public class Order
{
    public int Id { get; set; }
    public string Customer { get; set; } = string.Empty;
    public List<Service> Services { get; set; } = new();
}

// Service class
public class Service
{
    public string Name { get; set; } = string.Empty;
}
