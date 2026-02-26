using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the source for the report.
    public class Person
    {
        // Made nullable to satisfy the non‑nullable warning in C# 8+.
        public string? Name { get; set; }
        public int Age { get; set; }

        // Example of a calculated property that can be used in a conditional block.
        public bool IsAdult => Age >= 18;
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare the data source – a collection of Person objects.
            // -----------------------------------------------------------------
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice",   Age = 23 },
                new Person { Name = "Bob",     Age = 17 },
                new Person { Name = "Charlie", Age = 34 }
            };

            // The anonymous object wraps the collection so it can be referenced
            // in the template as "ds.Persons".
            var dataSource = new { Persons = people };

            // -----------------------------------------------------------------
            // 2. Load the WORDML template.
            //    The template contains a default conditional block, e.g.:
            //    <<foreach [in ds.Persons]>>
            //        <<if [IsAdult]>>
            //            <<[Name]>> is an adult.
            //        <<else>>
            //            <<[Name]>> is a minor.
            //        <<endif>>
            //    <<endforeach>>
            // -----------------------------------------------------------------
            // LoadOptions is not required for WORDML – Aspose.Words detects the format automatically.
            Document template = new Document("Template.xml");

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            //    - AllowMissingMembers prevents exceptions if a field is absent.
            //    - RemoveEmptyParagraphs cleans up paragraphs that become empty.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers |
                          ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Optional: customize the message shown for missing members.
            engine.MissingMemberMessage = "N/A";

            // -----------------------------------------------------------------
            // 4. Build the report by merging the data source with the template.
            //    The data source name "ds" is used inside the template tags.
            // -----------------------------------------------------------------
            engine.BuildReport(template, dataSource, "ds");

            // -----------------------------------------------------------------
            // 5. Save the populated document.
            // -----------------------------------------------------------------
            template.Save("Report.docx", SaveFormat.Docx);
        }
    }
}
