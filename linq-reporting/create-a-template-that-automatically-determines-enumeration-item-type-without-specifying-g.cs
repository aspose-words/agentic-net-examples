using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingEnumAutoDetect
{
    // Simple data class used in the collection.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    // Wrapper model that holds a strongly‑typed collection.
    // The template does not need to know the generic type; it only iterates over Items.
    public class ReportModel
    {
        public List<Person> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title paragraph.
            builder.Writeln("People Report");
            builder.Writeln();

            // Begin foreach over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a simple table to display each person's data.
            var table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Age");
            builder.EndRow();

            // Data row – the engine will infer that 'item' is a Person at runtime.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Age]>>");
            builder.EndRow();

            // End table.
            builder.EndTable();

            // End foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data model.
            // -----------------------------------------------------------------
            var template = new Document(templatePath);

            var model = new ReportModel();
            model.Items.Add(new Person { Name = "Alice", Age = 30 });
            model.Items.Add(new Person { Name = "Bob", Age = 25 });
            model.Items.Add(new Person { Name = "Charlie", Age = 28 });

            // -----------------------------------------------------------------
            // 3. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);

            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
