using System;
using System.Data;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class MissingMemberLogger
{
    static void Main()
    {
        // Create a template document with references to missing members.
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[missingObject.First().Id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[Id]>><</foreach>>");

        // Configure the reporting engine to treat missing members as null literals.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            // Use a unique token that will replace each missing member occurrence.
            MissingMemberMessage = "[MISSING]"
        };

        // Build the report using an empty data source (no "missingObject" defined).
        engine.BuildReport(builder.Document, new DataSet(), "");

        // After building, search the document for the placeholder token.
        List<int> missingPositions = new List<int>();
        string docText = builder.Document.GetText();

        int index = 0;
        while ((index = docText.IndexOf("[MISSING]", index, StringComparison.Ordinal)) != -1)
        {
            missingPositions.Add(index);
            index += "[MISSING]".Length;
        }

        // Log the occurrences.
        Console.WriteLine($"Total missing member occurrences: {missingPositions.Count}");
        for (int i = 0; i < missingPositions.Count; i++)
        {
            Console.WriteLine($"Occurrence {i + 1}: character position {missingPositions[i]}");
        }

        // Save the resulting document (optional).
        builder.Document.Save("ReportWithMissingMembers.docx");
    }
}
