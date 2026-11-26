using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvRecord
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Comment { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create a sample CSV file that uses a custom single‑quote character
        // as the text qualifier (quote character).
        // -----------------------------------------------------------------
        const string csvFileName = "data.csv";
        const char customQuote = '\'';

        var csvContent = new StringBuilder();
        // Header row.
        csvContent.AppendLine($"{customQuote}Id{customQuote},{customQuote}Name{customQuote},{customQuote}Comment{customQuote}");
        // Data rows – inner double quotes are part of the text and should be kept.
        csvContent.AppendLine($"{customQuote}1{customQuote},{customQuote}Alice{customQuote},{customQuote}She said \"\"Hello\"\" to me{customQuote}");
        csvContent.AppendLine($"{customQuote}2{customQuote},{customQuote}Bob{customQuote},{customQuote}He replied \"\"Hi!\"\"{customQuote}");

        File.WriteAllText(csvFileName, csvContent.ToString(), Encoding.UTF8);

        // -----------------------------------------------------------------
        // Build a LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        const string templateFileName = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("CSV Report");
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Comment: <<[row.Comment]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // -----------------------------------------------------------------
        // Load the CSV file and convert it into a strongly‑typed list.
        // -----------------------------------------------------------------
        var records = new List<CsvRecord>();
        var lines = File.ReadAllLines(csvFileName);
        // Skip header (first line).
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length != 3) continue;

            var idPart = parts[0].Trim(customQuote);
            var namePart = parts[1].Trim(customQuote);
            var commentPart = parts[2].Trim(customQuote);

            if (!int.TryParse(idPart, out int id)) continue;

            records.Add(new CsvRecord
            {
                Id = id,
                Name = namePart,
                Comment = commentPart
            });
        }

        // -----------------------------------------------------------------
        // Generate the report using the strongly‑typed data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templateFileName);
        var engine = new ReportingEngine();

        // Optional: enable reflection optimization for better performance.
        ReportingEngine.UseReflectionOptimization = true;

        engine.BuildReport(reportDoc, records, "Data");

        // Save the final document.
        const string outputFileName = "report.docx";
        reportDoc.Save(outputFileName);
    }
}
