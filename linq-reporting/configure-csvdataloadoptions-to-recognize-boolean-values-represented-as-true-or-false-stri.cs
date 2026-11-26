using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with boolean values as "true"/"false".
        string csvFile = "sample.csv";
        File.WriteAllText(csvFile, "Name,IsActive\nAlice,true\nBob,false");

        // Configure CSV loading options. The default parser already recognizes
        // "true" and "false" as boolean values, so no explicit BooleanTrueValues/
        // BooleanFalseValues properties are required.
        var csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
        };

        // Create a CSV data source using the configured options.
        var csvDataSource = new CsvDataSource(csvFile, csvOptions);

        // Build a simple Word template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags to iterate over CSV rows.
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Name: <<[item.Name]>> | Active: <<[item.IsActive]>>");
        builder.Writeln("<</foreach>>");

        // Generate the report using the CSV data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "items");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
