using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class DataModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a large JSON array.
        const int itemCount = 20000;
        var model = new DataModel();
        for (int i = 1; i <= itemCount; i++)
        {
            model.Persons.Add(new Person
            {
                Id = i,
                Name = $"Person {i}",
                Age = 20 + (i % 30)
            });
        }

        // Serialize to JSON file.
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(model));

        // Create LINQ Reporting template.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Id: <<[person.Id]>>, Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load data from JSON.
        var json = File.ReadAllText(jsonPath);
        var data = JsonConvert.DeserializeObject<DataModel>(json)!;

        // Benchmark without reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;
        var docWithoutOpt = new Document(templatePath);
        var engineWithoutOpt = new ReportingEngine();
        var swWithout = Stopwatch.StartNew();
        bool successWithout = engineWithoutOpt.BuildReport(docWithoutOpt, data, "data");
        swWithout.Stop();
        const string outputWithout = "report_without_opt.docx";
        docWithoutOpt.Save(outputWithout);
        Console.WriteLine($"Without optimization: {swWithout.ElapsedMilliseconds} ms, success = {successWithout}");

        // Benchmark with reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;
        var docWithOpt = new Document(templatePath);
        var engineWithOpt = new ReportingEngine();
        var swWith = Stopwatch.StartNew();
        bool successWith = engineWithOpt.BuildReport(docWithOpt, data, "data");
        swWith.Stop();
        const string outputWith = "report_with_opt.docx";
        docWithOpt.Save(outputWith);
        Console.WriteLine($"With optimization: {swWith.ElapsedMilliseconds} ms, success = {successWith}");
    }
}
