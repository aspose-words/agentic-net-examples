using System;
using Aspose.Words;
using Aspose.Words.Math;

class Program
{
    static void Main()
    {
        // Load the source document that contains OfficeMath equations.
        Document sourceDoc = new Document("OfficeMathSource.docx");

        // Retrieve all OfficeMath nodes in the document (including those in nested structures).
        NodeCollection officeMathNodes = sourceDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Create a new blank document that will hold the report.
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        // Write a header for the report.
        builder.Writeln("OfficeMath Report");
        builder.Writeln("-----------------");

        // Iterate through each OfficeMath node, output its index (position) and MathObjectType.
        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)officeMathNodes[i];
            builder.Writeln($"OfficeMath #{i + 1}: MathObjectType = {officeMath.MathObjectType}");
        }

        // Save the report document.
        reportDoc.Save("OfficeMathReport.docx");
    }
}
