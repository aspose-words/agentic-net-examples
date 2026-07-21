using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;

        public Item(int index, string name)
        {
            Index = index;
            Name = name;
        }
    }

    // Extension method container – used as a static helper in the template
    public static class MyExtensions
    {
        // Returns true if the item's Index is even
        public static bool IsEven(this Item item) => item.Index % 2 == 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Define file paths
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Report.docx");

            // ---------- Create template ----------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // LINQ Reporting tags
            builder.Writeln("<<foreach [item in Items]>>");
            // Use the static helper method syntax: MyExtensions.IsEven(item)
            builder.Writeln("<<if [MyExtensions.IsEven(item)]>>Even: <<[item.Name]>> <</if>>");
            builder.Writeln("<<if [!MyExtensions.IsEven(item)]>>Odd: <<[item.Name]>> <</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // ---------- Load template ----------
            Document reportDoc = new Document(templatePath);

            // ---------- Prepare data ----------
            ReportModel model = new ReportModel();
            model.Items.Add(new Item(1, "Alpha"));
            model.Items.Add(new Item(2, "Beta"));
            model.Items.Add(new Item(3, "Gamma"));
            model.Items.Add(new Item(4, "Delta"));

            // ---------- Build report ----------
            ReportingEngine engine = new ReportingEngine();
            // Register the static class that contains the extension method
            engine.KnownTypes.Add(typeof(MyExtensions));

            // Build the report using the root object name "model"
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            reportDoc.Save(resultPath);
        }
    }
}
