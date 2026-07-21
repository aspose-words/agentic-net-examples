using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

namespace AsposeWordsLinqReportingExample
{
    // Demonstrates LINQ Reporting with a CSV data source, a numbered list, filtering, and conditional formatting.
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV handling (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create sample CSV data.
            // -----------------------------------------------------------------
            const string csvPath = "data.csv";
            File.WriteAllText(csvPath,
                "Id,Name,Value\n" +
                "1,Alpha,5\n" +
                "2,Beta,12\n" +
                "3,Gamma,8\n" +
                "4,Delta,15\n" +
                "5,Epsilon,22");

            // -----------------------------------------------------------------
            // 2. Build a template document programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Create a numbered list style and apply it to the following paragraphs.
            List numberedList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;

            // Place <<restartNum>> and <<foreach>> in the same numbered paragraph.
            builder.Writeln("<<restartNum>><<foreach [row in persons]>>");

            // Filter rows where Value > 10.
            builder.Writeln("<<if [row.Value > 10]>>");

            // Conditional background color for even values.
            builder.Writeln(
                "<<if [row.Value % 2 == 0]>>" +
                "<<backColor [\"LightGray\"]>><<[row.Name]>> <</backColor>>" +
                "<<else>>" +
                "<<[row.Name]>>" +
                "<</if>> - <<[row.Value]>>");

            // End the outer if condition.
            builder.Writeln("<</if>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Remove list formatting after the loop.
            builder.ListFormat.RemoveNumbers();

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Configure CSV data source to treat the first line as headers.
            var loadOptions = new CsvDataLoadOptions(true);
            var csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            engine.BuildReport(reportDoc, csvDataSource, "persons");

            // Save the final report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
