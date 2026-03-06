using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document and attach a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert template expressions that reference a missing object member.
        // These expressions will be processed by ReportingEngine.
        builder.Writeln("<<[missingObject.First().Id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure ReportingEngine to allow missing members and define a custom message.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        engine.MissingMemberMessage = "Missing";

        // Build the report using an empty DataSet (no data for missingObject).
        engine.BuildReport(builder.Document, new DataSet(), "");

        // Add a custom document property.
        doc.CustomDocumentProperties.Add("AuthorName", "John Doe");

        // Insert a DOCPROPERTY field that reads the custom property.
        FieldDocProperty authorField = (FieldDocProperty)builder.InsertField(" DOCPROPERTY AuthorName ");
        authorField.Update();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
