using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Public properties with getters and setters.
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple Word template that uses only getters.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        // No setter calls are placed in the template.

        // -----------------------------------------------------------------
        // 2. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        // Aspose.Words ReportingEngine does not expose a RestrictedMembers collection.
        // To prevent template misuse, simply avoid exposing setter calls in the template.
        // The engine will only evaluate the expressions that are present.
        ReportingEngine engine = new ReportingEngine();

        // -----------------------------------------------------------------
        // 3. Build the report using a populated data object.
        // -----------------------------------------------------------------
        Person data = new Person { Name = "Alice", Age = 28 };
        // The root object name "person" must match the name used in the template tags.
        engine.BuildReport(template, data, "person");

        // -----------------------------------------------------------------
        // 4. Save the generated document.
        // -----------------------------------------------------------------
        template.Save("RestrictedMembersReport.docx");
    }
}
