using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "AddressTemplate.docx");
        string reportPath = Path.Combine(outputDir, "AddressReport.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Address Information:");
        builder.Writeln();

        // Conditional block for PO Box address.
        builder.Writeln("<<if [model.Address.IsPoBox]>>");
        builder.Writeln("PO Box: <<[model.Address.PoBoxNumber]>>");
        builder.Writeln("<</if>>");

        // Conditional block for full street address.
        builder.Writeln("<<if [!model.Address.IsPoBox]>>");
        builder.Writeln("Street: <<[model.Address.Street]>>");
        builder.Writeln("City: <<[model.Address.City]>>");
        builder.Writeln("State: <<[model.Address.State]>>");
        builder.Writeln("ZIP: <<[model.Address.Zip]>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Address = new AddressInfo
            {
                // Change IsPoBox to true to test the PO Box branch.
                IsPoBox = false,
                PoBoxNumber = "PO Box 789",
                Street = "123 Main St",
                City = "Anytown",
                State = "NY",
                Zip = "12345"
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model definitions.
// ---------------------------------------------------------------------
public class ReportModel
{
    public AddressInfo Address { get; set; } = new();
}

public class AddressInfo
{
    public bool IsPoBox { get; set; }
    public string PoBoxNumber { get; set; } = "";
    public string Street { get; set; } = "";
    public string City { get; set; } = "";
    public string State { get; set; } = "";
    public string Zip { get; set; } = "";
}
