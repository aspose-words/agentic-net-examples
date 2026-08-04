using System;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a CSV file with European date format (dd/MM/yyyy).
        const string csvFile = "people.csv";
        File.WriteAllText(csvFile,
            "Name,Date\r\n" +
            "Alice,25/12/2022\r\n" +
            "Bob,01/01/2023");

        // Set the thread culture to a European culture (en-GB) – this will be used for parsing dates.
        CultureInfo europeanCulture = new CultureInfo("en-GB");
        System.Threading.Thread.CurrentThread.CurrentCulture = europeanCulture;

        // Configure CSV load options. The HasHeaders flag is true because the first line contains column names.
        var loadOptions = new CsvDataLoadOptions(true);
        // If the library version supports it, assign the culture used for type conversion.
        // This ensures dates like "25/12/2022" are interpreted correctly.
        // Uncomment the following line if the property exists in your version:
        // loadOptions.CultureInfo = europeanCulture;

        // Create the CSV data source.
        var csvDataSource = new CsvDataSource(csvFile, loadOptions);

        // Build a simple template document that iterates over the CSV rows.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Date: <<[p.Date]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // Save the generated report.
        doc.Save("PeopleReport.docx");
    }
}
