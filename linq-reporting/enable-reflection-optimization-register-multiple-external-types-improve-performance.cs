using System;
using System.Collections.Generic;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    public class Program
    {
        // Entry point required for a console application.
        public static void Main(string[] args)
        {
            // Configure the reporting engine.
            ReportingEngineSetup.Configure();

            // Here you could invoke the engine to build a report, e.g.:
            // var engine = new ReportingEngine();
            // engine.BuildReport("Template.docx", "Output.docx", dataSource);
            Console.WriteLine("Reporting engine configured successfully.");
        }
    }

    public static class ReportingEngineSetup
    {
        public static void Configure()
        {
            // Enable reflection optimization (default is true, set explicitly for clarity)
            ReportingEngine.UseReflectionOptimization = true;

            // Create a new ReportingEngine instance
            ReportingEngine engine = new ReportingEngine();

            // Register external types that will be referenced in report templates.
            // These types can be used for static member access, type casts, etc.
            engine.KnownTypes.Add(typeof(Person));
            engine.KnownTypes.Add(typeof(Address));
            engine.KnownTypes.Add(typeof(List<>)); // Register a generic type as an example

            // Optional: configure additional engine options, such as allowing missing members.
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            // The engine is now ready to be used with BuildReport methods elsewhere in the application.
        }
    }

    // Sample external types that might appear in hierarchical data structures.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public Address HomeAddress { get; set; } = new();
    }

    public class Address
    {
        public string Street { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
    }
}
