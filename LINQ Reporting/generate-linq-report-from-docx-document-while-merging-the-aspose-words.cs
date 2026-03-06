using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains ReportingEngine tags, e.g. <<[people.Name]>>.
        string templatePath = "Template.docx";

        // Load the template document (uses the Document(string) constructor rule).
        Document doc = new Document(templatePath);

        // Prepare a LINQ data source – a list of POCO objects.
        List<Person> data = GetSampleData();

        // Create the reporting engine and populate the template.
        // The overload BuildReport(Document, object, string) is used to give the data source a name.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "people");

        // Save the generated report (uses the Document.Save(string) rule).
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    // Simple POCO class that will be used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    // Generates sample data using LINQ.
    private static List<Person> GetSampleData()
    {
        string[] names = { "Alice", "Bob", "Charlie" };
        string[] cities = { "London", "Paris", "Berlin" };

        // LINQ query creates a collection of Person objects.
        var query = from i in Enumerable.Range(0, names.Length)
                    select new Person
                    {
                        Name = names[i],
                        Age = 20 + i * 5,
                        City = cities[i]
                    };

        return query.ToList();
    }
}

/*
 * ILMerge integration:
 * After building the project, run ILMerge as a post‑build step to combine
 * Aspose.Words.dll with the compiled executable (MyApp.exe) into a single
 * assembly (MergedApp.exe). Example command line:
 *
 * ilmerge /out:MergedApp.exe Aspose.Words.dll MyApp.exe
 *
 * The merged assembly can then be distributed as a single executable.
 */
