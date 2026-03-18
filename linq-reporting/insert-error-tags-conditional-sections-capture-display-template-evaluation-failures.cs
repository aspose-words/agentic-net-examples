using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a conditional section that references a member which does not exist.
        // When the engine cannot resolve the member, it will insert the value of
        // ReportingEngine.MissingMemberMessage (set below) into the output.
        builder.Writeln("<<if [missingObject]>>This text is shown only if the condition is true.<<endif>>");

        // Add some regular content to demonstrate that the rest of the document is processed normally.
        builder.Writeln("Static content that is always rendered.");

        // Configure the ReportingEngine:
        //   - InlineErrorMessages: forces the engine to embed any template syntax errors directly into the output.
        //   - AllowMissingMembers: prevents the engine from throwing an exception when a member is missing.
        //   - MissingMemberMessage: the text that will replace a missing member reference.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages | ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "<<error>>"
        };

        // Build the report using an empty data source (no "missingObject" defined).
        // The conditional section will fail to evaluate, and the engine will insert "<<error>>"
        // where the missing member reference occurs.
        engine.BuildReport(doc, new DataSet());

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
