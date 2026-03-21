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

        // Prepare the data source. Use a visible type instead of an anonymous object.
        var data = new ReportData
        {
            customer = new Customer { 名前 = "山田太郎" } // Literal Japanese characters.
        };

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Message to display when a member is missing (optional).
        engine.MissingMemberMessage = "N/A";

        // Build the report using the template and the data source.
        engine.BuildReport(doc, data, "Data");

        // Save the generated document.
        doc.Save("ReportWithUnicode.docx");
    }
}

// Wrapper class for the data source (must be a visible type).
public class ReportData
{
    public Customer customer { get; set; }
}

// Class that contains a Unicode property name.
// The identifier is written directly, without any \uXXXX escape sequences.
public class Customer
{
    public string 名前 { get; set; } // Literal Japanese characters as the property name.
}
