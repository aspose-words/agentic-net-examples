using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CompositeReportExample
{
    public static void Main()
    {
        // Register encoding provider for CSV parsing (required for code pages).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create sample XML data file.
        // -----------------------------------------------------------------
        string xmlPath = "people.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Persons>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
    <Person>
        <Name>Bob Johnson</Name>
        <Age>40</Age>
    </Person>
</Persons>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Create sample CSV data file.
        // -----------------------------------------------------------------
        string csvPath = "orders.csv";
        string[] csvLines =
        {
            "OrderId,Product,Quantity",   // Header line – required for column names.
            "1001,Apple,5",
            "1002,Banana,12",
            "1003,Orange,7"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Build the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("=== People Report ===");
        builder.Writeln("");
        builder.Writeln("<<foreach [person in xml]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("");
        builder.Writeln("=== Orders Report ===");
        builder.Writeln("");
        builder.Writeln("<<foreach [order in csv]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Product: <<[order.Product]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // Load data sources.
        // -----------------------------------------------------------------
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // CSV data source – specify that the first line contains headers.
        CsvDataLoadOptions csvLoadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, csvLoadOptions);

        // -----------------------------------------------------------------
        // Build the report using both data sources.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        bool success = engine.BuildReport(
            doc,
            new object[] { xmlDataSource, csvDataSource },
            new string[] { "xml", "csv" });

        // Optional: check success flag (useful when InlineErrorMessages option is set).
        if (!success)
        {
            Console.WriteLine("Report generation encountered errors.");
        }

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        doc.Save("CompositeReport.docx");
    }
}
