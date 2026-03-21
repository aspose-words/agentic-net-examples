using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        const string sourcePath = "OfficeMathSource.docx";
        Document sourceDoc;

        if (File.Exists(sourcePath))
        {
            sourceDoc = new Document(sourcePath);
        }
        else
        {
            // Create a minimal source document if the expected file is missing.
            sourceDoc = new Document();
            var builder = new DocumentBuilder(sourceDoc);
            builder.Writeln("Source document not found. This placeholder document contains no OfficeMath equations.");
        }

        // Retrieve all OfficeMath nodes in the document (including those in nested structures).
        NodeCollection officeMathNodes = sourceDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Create a new blank document that will hold the report.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        // Write a header for the report.
        reportBuilder.Writeln("OfficeMath Report");
        reportBuilder.Writeln("-----------------");

        // Iterate through each OfficeMath node, output its index (position) and MathObjectType.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];
            reportBuilder.Writeln($"OfficeMath #{i + 1}: MathObjectType = {officeMath.MathObjectType}");
        }

        // If no OfficeMath nodes were found, indicate that in the report.
        if (officeMathNodes.Count == 0)
        {
            reportBuilder.Writeln("No OfficeMath equations were found in the source document.");
        }

        // Save the report document.
        reportDoc.Save("OfficeMathReport.docx");
    }
}
