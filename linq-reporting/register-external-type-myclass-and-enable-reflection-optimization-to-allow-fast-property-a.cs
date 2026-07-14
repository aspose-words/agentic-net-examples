using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MyClass
{
    // Static property that will be accessed from the template.
    public static string Greeting => "Hello from MyClass!";
}

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for faster property access.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a template document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert a LINQ Reporting tag that references the external type.
        builder.Writeln("<<[MyClass.Greeting]>>");

        // Set up the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the external type so it can be used in the template.
        engine.KnownTypes.Add(typeof(MyClass));

        // Build the report. A dummy data source is sufficient because the template only uses the external type.
        object dummyData = new object();
        engine.BuildReport(doc, dummyData, "data");

        // Save the generated document.
        doc.Save("ReportOutput.docx");
    }
}
