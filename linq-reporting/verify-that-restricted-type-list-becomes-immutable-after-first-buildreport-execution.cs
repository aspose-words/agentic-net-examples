using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class RestrictedTypesImmutabilityDemo
{
    static void Main()
    {
        // Create a simple template document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<[obj.ToString()]>>"); // placeholder, not used.

        // Set restricted types BEFORE the first BuildReport call.
        ReportingEngine.SetRestrictedTypes(new Type[] { typeof(System.Type) });

        // Verify that the restricted types were set correctly.
        Type[] initialRestricted = ReportingEngine.GetRestrictedTypes();
        Console.WriteLine("Initial restricted types count: " + initialRestricted.Length); // Should be 1

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        engine.BuildReport(doc, new object());

        // After the first BuildReport, attempting to modify the restricted types should throw.
        try
        {
            ReportingEngine.SetRestrictedTypes(new Type[] { typeof(string) });
            Console.WriteLine("ERROR: SetRestrictedTypes did not throw after BuildReport.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }

        // The array returned by GetRestrictedTypes is a copy; modifying it does not affect the engine.
        Type[] copy = ReportingEngine.GetRestrictedTypes();
        if (copy.Length > 0)
        {
            // Attempt to change the first element (won't affect the engine's internal list).
            copy[0] = typeof(string);
        }

        // Verify that the engine's restricted types remain unchanged.
        Type[] afterAttempt = ReportingEngine.GetRestrictedTypes();
        Console.WriteLine("Restricted types count after modification attempt: " + afterAttempt.Length);
        Console.WriteLine("First restricted type still System.Type: " + (afterAttempt[0] == typeof(System.Type)));
    }
}
