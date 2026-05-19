using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection "Persons".
        builder.Writeln("<<foreach [person in Persons]>>");

        // Create a bookmark whose name is the concatenation of FirstName and LastName.
        builder.Writeln("<<bookmark [person.FirstName + \" \" + person.LastName]>>");

        // The content of the bookmark – display the full name.
        builder.Writeln("<<[person.FirstName]>> <<[person.LastName]>>");

        // Close the bookmark and the foreach block.
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for report generation.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { FirstName = "John", LastName = "Doe" },
                new Person { FirstName = "Jane", LastName = "Smith" },
                new Person { FirstName = "Alice", LastName = "Johnson" }
            }
        };

        // -------------------------------------------------
        // 4. Build the report using Aspose.Words ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }

    // Root data object referenced in the template as "model".
    public class ReportModel
    {
        // Collection that will be iterated over in the template.
        public List<Person> Persons { get; set; } = new();
    }

    // Simple data entity with two fields to be concatenated.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
    }
}
