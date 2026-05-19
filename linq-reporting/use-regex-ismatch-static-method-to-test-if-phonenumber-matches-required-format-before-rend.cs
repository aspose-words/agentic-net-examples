using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportWrapper
        {
            Items = new List<ReportModel>
            {
                new ReportModel { PhoneNumber = "123-456-7890" }, // valid
                new ReportModel { PhoneNumber = "555-1234" }      // invalid
            }
        };

        // Create the template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Phone Numbers Report");
        builder.Writeln("---------------------");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item:");
        // Show the phone number if it matches the required format.
        builder.Writeln("<<if [Regex.IsMatch(item.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>");
        builder.Writeln("<<[item.PhoneNumber]>> (valid)");
        builder.Writeln("<</if>>");
        // Otherwise indicate that it is invalid.
        builder.Writeln("<<if [!Regex.IsMatch(item.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>");
        builder.Writeln("Invalid phone number");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Regex)); // Allow use of Regex static methods in the template.

        // Build the report.
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        templateDoc.Save("Report.docx");
    }
}

// Wrapper class that holds the collection referenced by the template.
public class ReportWrapper
{
    public List<ReportModel> Items { get; set; } = new();
}

// Simple data model with a phone number field.
public class ReportModel
{
    public string PhoneNumber { get; set; } = string.Empty;
}
