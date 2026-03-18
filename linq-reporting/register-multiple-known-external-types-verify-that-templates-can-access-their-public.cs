using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineKnownTypesDemo
{
    // Sample external types whose public members we want to expose to templates.
    public static class Utils
    {
        // Public static property that will be accessed from the template.
        public static string Greeting => "Hello from Utils!";
    }

    public class Person
    {
        // Public instance properties.
        public string FirstName { get; set; }
        public string LastName { get; set; }

        // Public method (also accessible if needed).
        public string FullName() => $"{FirstName} {LastName}";
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a simple template document containing placeholders that reference the known types.
            //    The LINQ Reporting Engine uses the syntax <<[TypeName.Member]>> for static members.
            //    For instance, <<[ReportingEngineKnownTypesDemo.Utils.Greeting]>> will be replaced with the value of Utils.Greeting.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Static greeting: <<[ReportingEngineKnownTypesDemo.Utils.Greeting]>>");
            builder.Writeln("Person full name: <<[person.FullName()]>>"); // will be filled from data source.

            // 2. Register the external types with the ReportingEngine's KnownTypes set.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Utils));   // static members
            engine.KnownTypes.Add(typeof(Person)); // instance members (used via data source)

            // 3. Prepare a data source that contains a Person instance.
            var data = new
            {
                person = new Person { FirstName = "John", LastName = "Doe" }
            };

            // 4. Build the report – the engine will replace the placeholders.
            engine.BuildReport(template, data);

            // 5. Verify that the static property was accessed without reflection.
            string resultText = template.GetText().Trim();
            Console.WriteLine("Resulting document text:");
            Console.WriteLine(resultText);

            // Expected output contains the static greeting and the person's full name.
            if (resultText.Contains(Utils.Greeting) && resultText.Contains("John Doe"))
                Console.WriteLine("Success: Known types accessed correctly.");
            else
                Console.WriteLine("Failure: Known types not accessed as expected.");

            // 6. Save the generated document (optional, demonstrates usage of save rule).
            template.Save("GeneratedReport.docx");
        }
    }
}
