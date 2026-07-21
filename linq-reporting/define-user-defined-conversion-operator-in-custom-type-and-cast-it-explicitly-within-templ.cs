using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Custom type with an explicit conversion operator to int.
    public class MyNumber
    {
        public int Value { get; }

        public MyNumber(int value) => Value = value;

        // Explicit cast from MyNumber to int.
        public static explicit operator int(MyNumber number) => number.Value;

        public override string ToString() => Value.ToString();
    }

    // Model class used as the data source for the report.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings; will be overwritten in Main.
        public MyNumber Custom { get; set; } = new MyNumber(0);
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that explicitly casts the custom type to int.
            // The expression "(int)model.Custom" uses the explicit conversion operator defined above.
            builder.Writeln("Custom value: <<[(int)model.Custom]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel
            {
                Custom = new MyNumber(42) // Example value.
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(report, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            report.Save(outputPath);
        }
    }
}
