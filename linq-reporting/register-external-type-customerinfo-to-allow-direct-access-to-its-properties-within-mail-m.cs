using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CustomerInfo
{
    // Static properties expose sample data that can be accessed directly from the template.
    public static string Name => "John Doe";
    public static int Age => 42;
    public static string Email => "john.doe@example.com";
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags that reference the static members of CustomerInfo.
        builder.Writeln("Customer Report");
        builder.Writeln("----------------");
        builder.Writeln("Name : <<[CustomerInfo.Name]>>");
        builder.Writeln("Age  : <<[CustomerInfo.Age]>>");
        builder.Writeln("Email: <<[CustomerInfo.Email]>>");

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the external type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(CustomerInfo));

        // Build the report. No data source is required because the template only uses static members.
        engine.BuildReport(doc, new object(), "");

        // Save the generated document.
        doc.Save("CustomerReport.docx");
    }
}
