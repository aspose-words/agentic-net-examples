using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportExample
{
    // Simple data entity.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Wrapper class that will be used as the data source for the report.
    public class ReportData
    {
        public List<Person> Items { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document in memory.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a title.
            builder.Writeln("People Report");
            builder.Writeln();

            // Insert a single‑row list template using ReportingEngine syntax.
            // The <<foreach>> tag iterates over the collection "Items".
            // Inside the loop we output the person's Name and Age as a list item.
            builder.Writeln("<<foreach [in Items]>>");
            builder.Writeln("• <<[Name]>> (Age: <<[Age]>>)");
            builder.Writeln("<</foreach>>");

            // 2. Prepare the data source.
            var data = new ReportData
            {
                Items = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob",   Age = 45 },
                    new Person { Name = "Carol", Age = 27 }
                }
            };

            // 3. Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(template, data);

            // 4. Save the resulting document.
            template.Save("PeopleReport.docx");
        }
    }
}
