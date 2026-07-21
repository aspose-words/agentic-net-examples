using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags that reference the external type.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The tags will access the static type CustomerInfo after it is registered with the engine.
        builder.Writeln("Customer Name: <<[CustomerInfo.Name]>>");
        builder.Writeln("Customer Age:  <<[CustomerInfo.Age]>>");

        // Prepare the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the external type so its members can be used directly in the template.
        engine.KnownTypes.Add(typeof(CustomerInfo));

        // Build the report. No data source object is required because the template only uses the registered type.
        engine.BuildReport(doc, new object());

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Sample external type whose properties will be accessed from the template.
public class CustomerInfo
{
    // Static properties are used because the template accesses the type directly.
    public static string Name { get; set; } = "John Doe";
    public static int Age { get; set; } = 42;
}
