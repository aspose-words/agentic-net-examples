using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The template defines a variable that obtains a System.Type instance.
        // Accessing members of a restricted type should cause the engine to throw.
        builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>>");
        builder.Writeln("<<[typeVar]>>");

        // Restrict the System.Type type before building the report.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        ReportingEngine engine = new ReportingEngine();

        try
        {
            // BuildReport will attempt to evaluate the template.
            // Because System.Type is restricted, an exception is expected.
            engine.BuildReport(doc, new object());
            Console.WriteLine("Test failed: no exception was thrown.");
        }
        catch (Exception ex)
        {
            // Expected path – the engine should reject access to the restricted type.
            Console.WriteLine($"Expected exception caught: {ex.GetType().Name}");
            Console.WriteLine($"Message: {ex.Message}");
        }
    }
}
