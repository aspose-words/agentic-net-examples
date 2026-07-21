using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            string templatePath = "Template.docx";
            string outputPath = "ReportOutput.docx";

            // -------------------------------------------------
            // Step 1: Create the LINQ Reporting template document
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Persons Report");
            builder.Writeln();

            // Begin foreach loop over Persons collection
            builder.Writeln("<<foreach [p in Persons]>>");

            // Conditional line: only write when Name is not empty
            builder.Writeln("<<if [p.Name != \"\"]>>Name: <<[p.Name]>> (Age: <<[p.Age]>>) <</if>>");

            // End foreach loop
            builder.Writeln("<</foreach>>");
            builder.Writeln();
            builder.Writeln("End of Report.");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Step 2: Load the template for report generation
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // Step 3: Prepare sample data
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "", Age = 25 },          // Empty name – should produce an empty paragraph
                    new Person { Name = "Bob", Age = 40 },
                    new Person { Name = null ?? string.Empty, Age = 22 } // Null handled as empty string
                }
            };

            // -------------------------------------------------
            // Step 4: Build the report with removal of empty paragraphs
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report; root object name is "model" as used in the template tags
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // Step 5: Save the final document
            // -------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
