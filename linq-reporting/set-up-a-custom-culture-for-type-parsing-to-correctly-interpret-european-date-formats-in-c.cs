using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with European date format (dd/MM/yyyy).
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Date,Value\r\n25/12/2023,100\r\n01/01/2024,200", Encoding.UTF8);

        // Load CSV data manually, parsing dates with a European culture.
        List<Record> records = LoadCsv(csvPath, new CultureInfo("fr-FR"));

        // Create a Word template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Date: <<[row.Date.ToString(\"dd MMM yyyy\")]>>   Value: <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the list of records as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, records, "Data");

        // Save the generated report.
        string outputPath = "Report.docx";
        template.Save(outputPath);
    }

    private static List<Record> LoadCsv(string path, CultureInfo culture)
    {
        var list = new List<Record>();
        string[] lines = File.ReadAllLines(path);
        if (lines.Length <= 1)
            return list; // No data.

        // Skip header line.
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length != 2)
                continue;

            // Parse date using the supplied culture (expects dd/MM/yyyy).
            if (!DateTime.TryParseExact(parts[0].Trim(), "dd/MM/yyyy", culture, DateTimeStyles.None, out DateTime date))
                continue;

            if (!int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
                continue;

            list.Add(new Record { Date = date, Value = value });
        }

        return list;
    }

    public class Record
    {
        public DateTime Date { get; set; }
        public int Value { get; set; }
    }
}
