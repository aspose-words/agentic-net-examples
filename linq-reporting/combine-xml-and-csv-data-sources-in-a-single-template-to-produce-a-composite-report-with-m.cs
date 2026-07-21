using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CompositeReportExample
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data files (XML and CSV) in a temporary folder.
        // -----------------------------------------------------------------
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        // XML file containing a list of Person elements.
        string xmlPath = Path.Combine(dataFolder, "people.xml");
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

        // CSV file containing a list of products (header row + data rows).
        string csvPath = Path.Combine(dataFolder, "products.csv");
        File.WriteAllText(csvPath,
            "Name,Price\n" +
            "Apple,0.99\n" +
            "Banana,0.59\n" +
            "Cherry,2.49");

        // ---------------------------------------------------------------
        // 2. Create the Word template with LINQ Reporting tags.
        // ---------------------------------------------------------------
        string templatePath = Path.Combine(dataFolder, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // XML Persons section.
        builder.Writeln("=== XML Persons ===");
        // When an XmlDataSource represents a collection (the root element contains a list),
        // the collection itself is referenced directly – no extra property name.
        builder.Writeln("<<foreach [p in xml]>>");
        builder.Writeln("- <<[p.Name]>> is <<[p.Age]>> years old");
        builder.Writeln("<</foreach>>");

        builder.Writeln(); // empty line between sections

        // CSV Products section.
        builder.Writeln("=== CSV Products ===");
        builder.Writeln("<<foreach [prod in csv]>>");
        builder.Writeln("- <<[prod.Name]>> costs $<<[prod.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template for report generation.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create data source objects.
        XmlDataSource xmlSource = new XmlDataSource(xmlPath);

        // CSV options: first line contains headers.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvSource = new CsvDataSource(csvPath, csvOptions);

        // ---------------------------------------------------------------
        // 4. Build the report using both data sources.
        // ---------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options

        // BuildReport overload that accepts multiple data sources.
        bool success = engine.BuildReport(
            reportDoc,
            new object[] { xmlSource, csvSource },
            new string[] { "xml", "csv" });

        // ---------------------------------------------------------------
        // 5. Save the generated report.
        // ---------------------------------------------------------------
        string outputPath = Path.Combine(dataFolder, "CompositeReport.docx");
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success
            ? $"Report generated successfully: {outputPath}"
            : "Report generation failed.");
    }
}
