using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingKnownTypesExample
{
    // External type A
    public class ExternalTypeA
    {
        public string Name { get; set; } = "DefaultA";
        public int Id { get; set; } = 0;
    }

    // External type B
    public class ExternalTypeB
    {
        public double Value { get; set; } = 0.0;
    }

    // Root model for the report
    public class ReportModel
    {
        public ExternalTypeA TypeA { get; set; } = new();
        public ExternalTypeB TypeB { get; set; } = new();
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                TypeA = new ExternalTypeA { Name = "Alpha", Id = 123 },
                TypeB = new ExternalTypeB { Value = 456.78 },
                Items = new List<string> { "Item1", "Item2", "Item3" }
            };

            // Create template document programmatically
            string templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("External Type A Name: <<[model.TypeA.Name]>>");
            builder.Writeln("External Type A Id: <<[model.TypeA.Id]>>");
            builder.Writeln("External Type B Value: <<[model.TypeB.Value]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("- <<[item]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Create the reporting engine and register external types
            var engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(ExternalTypeA));
            engine.KnownTypes.Add(typeof(ExternalTypeB));

            // Build the report
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine("output", "ReportResult.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            reportDoc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
