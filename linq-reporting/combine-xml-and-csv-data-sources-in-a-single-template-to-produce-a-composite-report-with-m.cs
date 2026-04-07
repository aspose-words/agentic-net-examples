using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // Ensure reporting namespace is included

public class CompositeReportExample
{
    public static void Main()
    {
        // Prepare working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Define file paths.
        string xmlPath = Path.Combine(workDir, "People.xml");
        string csvPath = Path.Combine(workDir, "Products.csv");
        string templatePath = Path.Combine(workDir, "Template.docx");
        string outputPath = Path.Combine(workDir, "CompositeReport.docx");

        // 1. Create sample XML data.
        File.WriteAllText(xmlPath,
@"<People>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
</People>");

        // 2. Create sample CSV data (headers + rows).
        File.WriteAllText(csvPath,
@"Product,Price
Apple,1.20
Banana,0.80
Orange,1.50");

        // 3. Build the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Composite Report ===");
        builder.Writeln();

        // XML data section.
        builder.Writeln("People (from XML):");
        builder.Writeln("<<foreach [p in xml]>>");
        builder.Writeln("- <<[p.Name]>> (Age <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // CSV data section.
        builder.Writeln("Products (from CSV):");
        builder.Writeln("<<foreach [c in csv]>>");
        builder.Writeln("- <<[c.Product]>> : $<<[c.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 4. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 5. Create data source objects.
        XmlDataSource xmlSource = new XmlDataSource(xmlPath);

        // Configure CSV options to indicate that the first row contains headers.
        CsvDataLoadOptions csvLoadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvSource = new CsvDataSource(csvPath, csvLoadOptions);

        // 6. Build the report using both data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(reportDoc,
            new object[] { xmlSource, csvSource },
            new string[] { "xml", "csv" });

        // 7. Save the final report.
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
