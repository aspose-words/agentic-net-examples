using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class ContextualMemberAccessExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a built‑in document property that we will reference from the template.
        doc.BuiltInDocumentProperties.Title = "Sample Title";

        // Add a custom document property.
        doc.CustomDocumentProperties.Add("MyProp", "Custom value");

        // Insert a template expression that references an existing property.
        // The expression will be replaced with the value of the built‑in Title property.
        builder.Writeln("Document title: <<[doc.BuiltInDocumentProperties.Title]>>");

        // Insert a template expression that references a missing member.
        // Because we will enable AllowMissingMembers, the MissingMemberMessage will be printed.
        builder.Writeln("Missing member test: <<[missingObject.Name]>>");

        // Insert a DOCPROPERTY field that displays the custom property we added.
        FieldDocProperty customField = (FieldDocProperty)builder.InsertField(" DOCPROPERTY \"MyProp\" ");
        customField.Update();

        // Build the report using ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members so the engine does not throw an exception.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Text to display when a member is missing.
            MissingMemberMessage = "Missed"
        };

        // The data source is empty because we only use the document itself for the example.
        engine.BuildReport(builder.Document, new DataSet(), "");

        // Save the resulting document.
        doc.Save("ContextualMemberAccess.docx");
    }
}
