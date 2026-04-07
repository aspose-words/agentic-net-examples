using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupingExample
{
    // Simple data model representing a person.
    public class Person
    {
        public Person(string name, int age, string department)
        {
            Name = name;
            Age = age;
            Department = department;
        }

        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            List<Person> persons = new()
            {
                new Person("Alice Johnson", 28, "Human Resources"),
                new Person("Bob Smith", 35, "Finance"),
                new Person("Carol White", 42, "Human Resources"),
                new Person("David Brown", 31, "IT"),
                new Person("Eve Davis", 27, "Finance")
            };

            // 2. Create a template document programmatically.
            Document template = new();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Employees grouped by Department");
            builder.Writeln();

            // Outer foreach iterates over groups created by LINQ GroupBy.
            builder.Writeln("<<foreach [g in persons.GroupBy(p => p.Department)]>>");
            builder.Writeln("Department: <<[g.Key]>>");
            builder.Writeln();

            // Inner foreach iterates over persons within the current group.
            builder.Writeln("<<foreach [p in g]>>");
            builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln(); // Add a blank line between departments.
            builder.Writeln("<</foreach>>");

            // 3. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.Options = ReportBuildOptions.None;

            // Pass the collection as the data source and give it the name "persons"
            // so that the template can reference it directly.
            engine.BuildReport(template, persons, "persons");

            // 4. Save the generated report.
            const string outputPath = "GroupedReport.docx";
            template.Save(outputPath);
        }
    }
}
