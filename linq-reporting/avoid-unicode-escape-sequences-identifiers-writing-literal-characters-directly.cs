using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a reporting template that references a Unicode identifier directly.
        // The property name "名前" (Japanese for "Name") is written as literal characters.
        builder.Writeln("Customer name: <<[customer.名前]>>");

        // Prepare the data source. The anonymous object contains a "customer" object
        // whose type has a property named "名前".
        var data = new
        {
            customer = new Customer { 名前 = "山田太郎" } // Literal Japanese characters.
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members to be treated as null (optional for this example).
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Message to display when a member is missing (not used here but shown for completeness).
        engine.MissingMemberMessage = "N/A";

        // Build the report using the template and the data source.
        engine.BuildReport(doc, data, string.Empty);

        // Save the generated document.
        doc.Save("ReportWithUnicode.docx");
    }
}

// Class that contains a Unicode property name.
// The identifier is written directly, without any \uXXXX escape sequences.
public class Customer
{
    public string 名前 { get; set; } // Literal Japanese characters as the property name.
}
