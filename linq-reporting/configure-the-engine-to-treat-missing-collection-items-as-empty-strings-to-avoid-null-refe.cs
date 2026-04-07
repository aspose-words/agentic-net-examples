using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection "Items".
        builder.Writeln("<<foreach [item in Items]>>");
        // Output each item; if the item is missing (null), we want an empty string.
        builder.Writeln("Item: <<[item]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (required before BuildReport).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model with a collection that contains a null.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<string?> { "First", null, "Third" }
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine.
        //    - AllowMissingMembers treats missing members as null.
        //    - MissingMemberMessage set to empty string makes null appear as "".
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = string.Empty
        };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the template.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Collection of strings; some entries may be null.
    public List<string?> Items { get; set; } = new();
}
