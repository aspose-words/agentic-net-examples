using System;
using Aspose.Words.Reporting;

// Apply the attribute at the assembly level – must be placed before any other code (except using/extern)
[assembly: LoggingAssemblyAttribute("MyCustomLogging.dll")]

// Define a custom assembly attribute that specifies the logging assembly to be used
[AttributeUsage(AttributeTargets.Assembly, Inherited = false, AllowMultiple = false)]
public sealed class LoggingAssemblyAttribute : Attribute
{
    // Path or name of the logging assembly
    public string AssemblyName { get; }

    public LoggingAssemblyAttribute(string assemblyName)
    {
        AssemblyName = assemblyName;
    }
}

namespace ReportingEngineWithLogging
{
    class Program
    {
        static void Main()
        {
            // Create a new instance of the ReportingEngine
            ReportingEngine engine = new ReportingEngine();

            // Example: set options to allow missing members and inline error messages
            engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages;

            // Set a custom missing member message
            engine.MissingMemberMessage = "Missing data";

            // Normally you would load a template document and build a report here
            // Document doc = new Document("Template.docx");
            // engine.BuildReport(doc, dataSource);

            // For demonstration purposes we just output that the engine is configured
            Console.WriteLine("ReportingEngine configured with custom logging assembly: MyCustomLogging.dll");
        }
    }
}
