using System;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register types from the core .NET assemblies.
        engine.KnownTypes.Add(typeof(System.Math));               // mscorlib/System.Runtime
        engine.KnownTypes.Add(typeof(System.Net.WebClient));      // System.Net

        // Load a custom assembly at runtime and register a type from it.
        // Ensure that "MyCustomLib.dll" is accessible at the specified path.
        Assembly customAssembly = Assembly.LoadFrom("MyCustomLib.dll");
        Type customType = customAssembly.GetType("MyCustomLib.MyHelper");
        if (customType != null)
            engine.KnownTypes.Add(customType);

        // Load the template document that contains LINQ Reporting Engine tags.
        Document doc = new Document("Template.docx");

        // Build the report. An empty anonymous object is used as a placeholder data source.
        engine.BuildReport(doc, new { });

        // Save the populated document.
        doc.Save("Report.docx");
    }
}
