using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Data model for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
        public int MinAge { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // ---------- Create the template document ----------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Header showing the runtime filter value.
            builder.Writeln("Filtered Persons (Age >= <<[model.MinAge]>>):");

            // Loop over the collection and apply a runtime filter using an IF tag.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("<<if [p.Age >= model.MinAge]>>");
            builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
            builder.Writeln("<</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // ---------- Load the template and build the report ----------
            Document reportDoc = new Document(templatePath);

            // Sample data with a dynamic filter criterion.
            ReportModel model = new ReportModel
            {
                MinAge = 30,
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 25 },
                    new Person { Name = "Bob", Age = 35 },
                    new Person { Name = "Charlie", Age = 40 }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
