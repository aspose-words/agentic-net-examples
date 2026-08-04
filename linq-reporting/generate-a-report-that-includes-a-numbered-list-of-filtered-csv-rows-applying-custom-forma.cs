using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample CSV file.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllText(csvPath,
            "Id,Name,Value\r\n" +
            "1,Alpha,30\r\n" +
            "2,Beta,60\r\n" +
            "3,Gamma,45\r\n" +
            "4,Delta,80\r\n");

        // 2. Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a numbered list style and apply it to subsequent paragraphs.
        List numberedList = template.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Insert LINQ Reporting tags.
        // The <<restartNum>> tag must be placed immediately before <<foreach>> in the same numbered paragraph.
        builder.Writeln("<<restartNum>><<foreach [row in csv]>>");

        // Conditional formatting: rows with Value > 50 get a light gray background.
        builder.Writeln(
            "<<if [row.Value > 50]>>" +
            "<<backColor [\"LightGray\"]>><<[row.Name]>> <</backColor>><</if>>" +
            "<<if [row.Value <= 50]>>" +
            "<<[row.Name]>>" +
            "<</if>>");

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, just for inspection).
        string templatePath = Path.Combine(workDir, "template.docx");
        template.Save(templatePath);

        // 3. Prepare CSV data source with load options.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions
        {
            HasHeaders = true,
            Delimiter = ',',
            QuoteChar = '"',
            CommentChar = '#'
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // The data source name used in the template tags is "csv".
        engine.BuildReport(template, csvData, "csv");

        // 5. Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        template.Save(reportPath);

        Console.WriteLine("Report generated at: " + reportPath);
    }
}
