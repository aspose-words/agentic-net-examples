using System;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words.Reporting;

namespace ReportingEngineExample
{
    // Sample data class whose members will be inspected.
    public class MyData
    {
        public string Name { get; set; }
        public int Age { get; set; }
        // Non‑public member – will not be added to the list.
        private string Secret { get; set; }
    }

    public class Program
    {
        // List that will hold the names of public properties to be treated as restricted.
        private static readonly List<string> RestrictedMembers = new List<string>();

        public static void Main()
        {
            // Populate the RestrictedMembers list with the names of all public instance properties
            // of the type that will be used in the reporting engine.
            Type targetType = typeof(MyData);
            PropertyInfo[] publicProperties = targetType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in publicProperties)
            {
                RestrictedMembers.Add(prop.Name);
            }

            // Example output (optional, can be removed in production).
            Console.WriteLine("Restricted members for type " + targetType.FullName + ":");
            foreach (string memberName in RestrictedMembers)
            {
                Console.WriteLine("- " + memberName);
            }

            // Register the type as restricted before any report is built.
            // This prevents the reporting engine from accessing any members (including the public properties)
            // of MyData through template syntax.
            ReportingEngine.SetRestrictedTypes(targetType);

            // At this point you would normally build a report using ReportingEngine.
            // The engine will now treat accesses to MyData members as restricted.
            // Example (commented out because it requires a template document):
            // Document doc = new Document("Template.docx");
            // ReportingEngine engine = new ReportingEngine();
            // engine.BuildReport(doc, new MyData { Name = "John", Age = 30 });
        }
    }
}
