using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingUnicodeDemo
{
    // Model class with Unicode characters directly in property names.
    public class Person
    {
        // Property name "名前" (Japanese for "Name").
        public string 名前 { get; set; } = string.Empty;

        // Property name "年齢" (Japanese for "Age").
        public int 年齢 { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // Step 1: Create a Word template with LINQ tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Simple paragraph with Unicode property references.
            builder.Writeln("顧客情報:"); // "Customer Information:" in Japanese.
            builder.Writeln("名前: <<[model.名前]>>");
            builder.Writeln("年齢: <<[model.年齢]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template and prepare data.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data.
            Person person = new Person
            {
                名前 = "山田太郎", // Example Japanese name.
                年齢 = 30
            };

            // -------------------------------------------------
            // Step 3: Build the report using ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(reportDoc, person, "model");

            // -------------------------------------------------
            // Step 4: Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);

            // Indicate completion (no interactive input).
            Console.WriteLine("Report generated successfully: " + reportPath);
        }
    }
}
