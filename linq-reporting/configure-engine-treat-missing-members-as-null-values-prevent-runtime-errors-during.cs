using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a builder to insert template tags.
            DocumentBuilder builder = new DocumentBuilder();
            // Template contains a reference to a missing object member.
            builder.Writeln("<<[missingObject.First().Id]>>");
            builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

            // Configure the ReportingEngine to treat missing members as null.
            ReportingEngine engine = new ReportingEngine
            {
                // AllowMissingMembers makes the engine replace missing members with null.
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Optional: specify a placeholder text that will be printed for missing members.
            engine.MissingMemberMessage = "Missed";

            // Build the report using an empty data source (DataSet) to avoid real data.
            // The third parameter (data source name) can be an empty string when not needed.
            bool success = engine.BuildReport(builder.Document, new DataSet(), "");

            // The report is now built; you can save or further process the document as needed.
            // For demonstration, we simply output the result status.
            Console.WriteLine($"Report built successfully: {success}");
        }
    }
}
