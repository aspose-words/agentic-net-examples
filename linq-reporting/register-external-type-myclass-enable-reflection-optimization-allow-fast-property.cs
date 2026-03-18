using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MyClass
{
    public string Name { get; set; }
    public int Value { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for faster property access in templates.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a ReportingEngine instance.
        ReportingEngine engine = new ReportingEngine();

        // Register the external type so it can be used inside templates.
        engine.KnownTypes.Add(typeof(MyClass));

        // Prepare a simple (empty) template document.
        Document doc = new Document();

        // Build the report using an instance of the registered type.
        MyClass data = new MyClass { Name = "Sample", Value = 123 };
        engine.BuildReport(doc, data);

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
