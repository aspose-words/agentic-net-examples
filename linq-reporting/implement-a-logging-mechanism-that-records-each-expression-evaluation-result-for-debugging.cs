using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Central log that records every property evaluation during report generation.
    private static readonly List<string> EvaluationLog = new();

    public static void Main()
    {
        // Ensure the code page provider is available (required by Aspose.Words for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create the LINQ Reporting template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Template uses a foreach loop over Items and prints Index and Name.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // 2. Prepare realistic sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // 3. Build the report using Aspose.Words LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this example.
        engine.BuildReport(template, model, "model");

        // 4. Save the generated report.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string reportPath = Path.Combine(outputDir, "ReportOutput.docx");
        template.Save(reportPath);

        // 5. Write the evaluation log to a text file and also output to console.
        string logPath = Path.Combine(outputDir, "EvaluationLog.txt");
        File.WriteAllLines(logPath, EvaluationLog);
        Console.WriteLine("Report generated at: " + reportPath);
        Console.WriteLine("Evaluation log written to: " + logPath);
        Console.WriteLine("=== Evaluation Log ===");
        foreach (string entry in EvaluationLog)
        {
            Console.WriteLine(entry);
        }
    }

    // Root data model for the report.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    // Data item whose property getters log each access.
    public class Item
    {
        private int _index;
        private string _name = string.Empty;

        public int Index
        {
            get => Log(nameof(Index), _index);
            set => _index = value;
        }

        public string Name
        {
            get => Log(nameof(Name), _name);
            set => _name = value ?? string.Empty;
        }

        // Generic logging helper.
        private static T Log<T>(string propertyName, T value)
        {
            EvaluationLog.Add($"{propertyName} accessed, value: {value}");
            return value;
        }
    }
}
