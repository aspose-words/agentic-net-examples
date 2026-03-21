using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace MyCustomLib
{
    public class MyHelper
    {
        public static double Square(double x) => x * x;
        public string Greeting => "Hello from MyHelper";
    }

    // Simple data source class required by ReportingEngine
    public class ReportData
    {
        public string Title { get; set; } = "Sample Report";
    }
}

class Program
{
    static void Main()
    {
        var engine = new ReportingEngine();

        // Register types from the core .NET assemblies.
        engine.KnownTypes.Add(typeof(System.Math));               // mscorlib/System.Runtime
        engine.KnownTypes.Add(typeof(System.Net.WebClient));      // System.Net

        // Register a custom type defined in this project.
        engine.KnownTypes.Add(typeof(MyCustomLib.MyHelper));

        // Create an empty document (or load a template if you have one).
        var doc = new Document();

        // Use a visible data source type.
        var data = new MyCustomLib.ReportData();

        // Build the report.
        engine.BuildReport(doc, data);

        // Save the populated document.
        doc.Save("Report.docx");
    }
}
