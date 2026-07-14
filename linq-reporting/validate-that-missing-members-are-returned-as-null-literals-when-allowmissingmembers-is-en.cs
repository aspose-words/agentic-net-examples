using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MissingMembersDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags that reference members which do not exist.
        builder.Writeln("Member: <<[missingObject.Name]>>");
        builder.Writeln("Foreach test: <<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure the reporting engine to treat missing members as null literals.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Do not set MissingMemberMessage so that missing members are rendered as empty (null).

        // Build the report using an empty DataSet as the data source.
        engine.BuildReport(doc, new DataSet(), "");

        // Save the resulting document (optional, for visual inspection).
        const string outputPath = "ReportWithMissingMembers.docx";
        doc.Save(outputPath);

        // Retrieve the plain text of the document to verify the output.
        string resultText = doc.GetText();

        // Validation: missing members should produce empty placeholders, not the literal text "Missed" or any value.
        bool memberIsEmpty = resultText.Contains("Member: ") && !resultText.Contains("Member: Missed") && !resultText.Contains("Member: null");
        bool foreachIsEmpty = resultText.Contains("Foreach test: ") && !resultText.Contains("Foreach test: null");

        if (memberIsEmpty && foreachIsEmpty)
        {
            Console.WriteLine("Missing members were rendered as null (empty) as expected.");
        }
        else
        {
            Console.WriteLine("Unexpected output. Verify that AllowMissingMembers is enabled correctly.");
        }
    }
}
