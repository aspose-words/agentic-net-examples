using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CompositeReportExample
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlPath = "employees.xml";
        File.WriteAllText(xmlPath,
@"<Employees>
    <Employee>
        <Name>John Doe</Name>
        <Department>Finance</Department>
    </Employee>
    <Employee>
        <Name>Jane Smith</Name>
        <Department>HR</Department>
    </Employee>
</Employees>");

        // Prepare sample CSV data.
        string csvPath = "sales.csv";
        File.WriteAllText(csvPath,
@"Product,Quantity
Laptop,5
Smartphone,12
Tablet,7");

        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Section for XML data (employees).
        builder.Writeln("Employees:");
        builder.Writeln("<<foreach [emp in employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Dept: <<[emp.Department]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Section for CSV data (sales).
        builder.Writeln("Sales:");
        builder.Writeln("<<foreach [sale in sales]>>");
        builder.Writeln("Product: <<[sale.Product]>> - Qty: <<[sale.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Load data sources.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ',',
            HasHeaders = true
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using both data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template,
            new object[] { xmlData, csvData },
            new string[] { "employees", "sales" });

        // Save the generated report.
        string outputPath = "CompositeReport.docx";
        template.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
