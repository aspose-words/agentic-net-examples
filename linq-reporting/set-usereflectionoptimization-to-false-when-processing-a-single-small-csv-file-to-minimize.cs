using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for CSV encodings.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // File paths.
            string csvPath = Path.Combine(outputDir, "people.csv");
            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "report.docx");

            // 1. Create a small CSV file with headers.
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Age,City",
                "Alice,30,New York",
                "Bob,25,London",
                "Charlie,35,Sydney"
            });

            // 2. Build a template document containing LINQ Reporting tags.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln();

            builder.Writeln("<<foreach [person in persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("City: <<[person.City]>>");
            builder.Writeln("<</foreach>>");

            // Save the template (demonstrates load/save lifecycle).
            templateDoc.Save(templatePath);

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Configure CSV data source options (headers present).
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(hasHeaders: true);

            // 5. Create the CSV data source.
            CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

            // 6. Disable reflection optimization for this small data set.
            ReportingEngine.UseReflectionOptimization = false;

            // 7. Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, csvDataSource, "persons");

            // 8. Save the generated report.
            doc.Save(resultPath);
        }
    }
}
