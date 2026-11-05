using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create the LINQ Reporting template
        const string templatePath = "AddressTemplate.docx";
        CreateTemplate(templatePath);

        // Sample data: PO Box address
        var poBoxAddress = new Address
        {
            IsPoBox = true,
            PoBoxNumber = "PO Box 123",
            Street = "",
            City = "Springfield",
            State = "IL",
            Zip = "62704"
        };

        // Sample data: Full street address
        var fullAddress = new Address
        {
            IsPoBox = false,
            PoBoxNumber = "",
            Street = "742 Evergreen Terrace",
            City = "Springfield",
            State = "IL",
            Zip = "62704"
        };

        // Generate reports for both samples
        GenerateReport(templatePath, poBoxAddress, "Report_POBox.docx");
        GenerateReport(templatePath, fullAddress, "Report_FullAddress.docx");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Address Report");
        builder.Writeln();

        // Conditional block for PO Box address
        builder.Writeln("<<if [address.IsPoBox]>>");
        builder.Writeln("PO Box: <<[address.PoBoxNumber]>>");
        builder.Writeln("<</if>>");

        // Conditional block for full street address (when not a PO Box)
        builder.Writeln("<<if [!address.IsPoBox]>>");
        builder.Writeln("Street: <<[address.Street]>>");
        builder.Writeln("City: <<[address.City]>>");
        builder.Writeln("State: <<[address.State]>>");
        builder.Writeln("ZIP: <<[address.Zip]>>");
        builder.Writeln("<</if>>");

        doc.Save(path);
    }

    private static void GenerateReport(string templatePath, Address data, string outputPath)
    {
        // Load the template document
        var doc = new Document(templatePath);

        // Build the report using the address object as the root named "address"
        var engine = new ReportingEngine();
        engine.BuildReport(doc, data, "address");

        // Save the generated report
        doc.Save(outputPath);
    }
}

// Data model representing an address
public class Address
{
    public bool IsPoBox { get; set; }
    public string PoBoxNumber { get; set; } = "";
    public string Street { get; set; } = "";
    public string City { get; set; } = "";
    public string State { get; set; } = "";
    public string Zip { get; set; } = "";
}
