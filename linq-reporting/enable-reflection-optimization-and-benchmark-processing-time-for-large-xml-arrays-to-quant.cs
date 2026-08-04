using System;
using System.Diagnostics;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for files used in the example.
        const string templatePath = "Template.docx";
        const string xmlPath = "Data.xml";
        const string reportOptimizedPath = "ReportOptimized.docx";
        const string reportNonOptimizedPath = "ReportNonOptimized.docx";

        // Create a large XML file with many Person elements.
        const int itemCount = 5000;
        var xmlDoc = new XDocument(
            new XElement("persons",
                new XElement("person",
                    new XElement("Name", "Name0"),
                    new XElement("Age", 0))
            )
        );

        var personsElement = xmlDoc.Root!;
        personsElement.RemoveAll(); // Clear the placeholder.

        for (int i = 0; i < itemCount; i++)
        {
            personsElement.Add(
                new XElement("person",
                    new XElement("Name", $"Name{i}"),
                    new XElement("Age", i % 100))
            );
        }

        xmlDoc.Save(xmlPath);

        // Create a LINQ Reporting template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Benchmark without reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;
        var docWithoutOpt = new Document(templatePath);
        var dataSourceWithout = new XmlDataSource(xmlPath);
        var engineWithout = new ReportingEngine();
        var swWithout = Stopwatch.StartNew();
        engineWithout.BuildReport(docWithoutOpt, dataSourceWithout, "persons");
        swWithout.Stop();
        docWithoutOpt.Save(reportNonOptimizedPath);
        long timeWithoutOpt = swWithout.ElapsedMilliseconds;

        // Benchmark with reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;
        var docWithOpt = new Document(templatePath);
        var dataSourceWith = new XmlDataSource(xmlPath);
        var engineWith = new ReportingEngine();
        var swWith = Stopwatch.StartNew();
        engineWith.BuildReport(docWithOpt, dataSourceWith, "persons");
        swWith.Stop();
        docWithOpt.Save(reportOptimizedPath);
        long timeWithOpt = swWith.ElapsedMilliseconds;

        // Output the measured times.
        Console.WriteLine($"Processing time without reflection optimization: {timeWithoutOpt} ms");
        Console.WriteLine($"Processing time with reflection optimization:    {timeWithOpt} ms");
    }
}
