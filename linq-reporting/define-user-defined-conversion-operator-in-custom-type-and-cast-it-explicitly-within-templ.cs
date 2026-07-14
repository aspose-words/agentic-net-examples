using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConversionExample
{
    // Custom type with an explicit conversion operator to int.
    public class MyNumber
    {
        private readonly int _value;

        public MyNumber(int value) => _value = value;

        // Explicit conversion from MyNumber to int.
        public static explicit operator int(MyNumber number) => number._value;
    }

    // Model class used as the data source for the report.
    public class ReportModel
    {
        // Property of the custom type.
        public MyNumber Number { get; set; } = new MyNumber(0);
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with a LINQ Reporting tag that
            //    explicitly casts the custom type to int.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // The expression uses C# cast syntax: (int)model.Number
            builder.Writeln("Number: <<[(int)model.Number]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document (simulating a separate load step).
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Number = new MyNumber(42) // Example value.
            };

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words LINQ Reporting Engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
