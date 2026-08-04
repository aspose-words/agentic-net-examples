using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank document and a builder to insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The tag attempts to obtain the base type of System.String.
            // Without restrictions this would output "System.Object".
            builder.Writeln("<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");

            // Restrict access to System.Type members to prevent reflection‑based code execution.
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            // Configure the engine to treat missing members as null instead of throwing.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report using an empty data source (no root object needed).
            engine.BuildReport(doc, new object(), "");

            // Save the resulting document.
            doc.Save("RestrictedReport.docx");
        }
    }
}
