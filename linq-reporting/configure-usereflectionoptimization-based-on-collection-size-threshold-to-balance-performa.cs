using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    // Threshold at which we switch the reflection optimization on/off.
    private const int CollectionSizeThreshold = 5;

    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a simple foreach loop that will list persons.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // (Optional) Save the template to disk – useful for inspection.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Prepare sample data.
        Model model = new Model();
        model.Persons.Add(new Person { Name = "Alice", Age = 30 });
        model.Persons.Add(new Person { Name = "Bob", Age = 25 });
        model.Persons.Add(new Person { Name = "Charlie", Age = 35 });
        model.Persons.Add(new Person { Name = "Diana", Age = 28 });
        model.Persons.Add(new Person { Name = "Eve", Age = 22 });
        model.Persons.Add(new Person { Name = "Frank", Age = 40 });

        // 3. Configure reflection optimization based on collection size.
        // If the collection is larger than the threshold, enable optimization;
        // otherwise, disable it to reduce overhead for small collections.
        ReportingEngine.UseReflectionOptimization = model.Persons.Count > CollectionSizeThreshold;

        // 4. Build the report.
        ReportingEngine engine = new ReportingEngine();
        // Explicitly set options (none in this case) as required by the rules.
        engine.Options = ReportBuildOptions.None;

        // BuildReport expects the root object name to match the tags in the template.
        engine.BuildReport(template, model, "model");

        // 5. Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
