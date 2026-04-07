using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths
        string csvPath = Path.Combine(outputDir, "Data.csv");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Create a simple CSV file with headers
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        });

        // -------------------- Create template document --------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Set custom page margins (1 inch = 72 points)
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;

        // Title
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach block – each iteration will add a new table
        builder.Writeln("<<foreach [person in persons]>>");

        // Start a table that will be repeated for each person
        Table table = builder.StartTable();

        // Header row (static, appears once per table)
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.InsertCell();
        builder.Write("City");
        builder.EndRow();

        // Data row – values are taken from the current person
        builder.InsertCell();
        builder.Write("<<[person.Name]>>");
        builder.InsertCell();
        builder.Write("<<[person.Age]>>");
        builder.InsertCell();
        builder.Write("<<[person.City]>>");
        builder.EndRow();

        // Finish the table
        builder.EndTable();

        // End foreach block
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);
        // ------------------------------------------------------------------

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Create CSV data source with header row detection
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(reportDoc, dataSource, "persons");

        // Save the final report
        reportDoc.Save(reportPath);
    }
}
