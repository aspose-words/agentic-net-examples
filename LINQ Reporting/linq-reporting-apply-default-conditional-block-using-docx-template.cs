using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // 2. Write a default conditional block into the template.
            //    The block prints "Adult" if Age >= 18, otherwise "Minor".
            //    The data source object will be referenced by the name "person".
            builder.Writeln("<<if [person.Age >= 18]>>Adult: <<[person.Name]>>");
            builder.Writeln("<<else>>Minor: <<[person.Name]>>");
            builder.Writeln("<<endif>>");

            // 3. Prepare the data source.
            Person data = new Person
            {
                Name = "John Doe",
                Age = 21
            };

            // 4. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // Optional: allow missing members without throwing exceptions.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // BuildReport overload that lets us reference the data source object itself.
            engine.BuildReport(template, data, "person");

            // 5. Save the populated document.
            string outputPath = "ReportWithConditionalBlock.docx";
            template.Save(outputPath);
        }
    }
}
