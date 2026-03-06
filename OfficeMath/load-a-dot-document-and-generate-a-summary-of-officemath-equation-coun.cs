using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Math;

class OfficeMathSummaryGenerator
{
    static void Main()
    {
        // Load the DOT template document.
        // The Document(string) constructor automatically detects the format.
        Document templateDoc = new Document("Template.dotx");

        // Build a textual summary of OfficeMath equation counts per section.
        StringBuilder summaryBuilder = new StringBuilder();
        int sectionNumber = 1;

        foreach (Section section in templateDoc.Sections)
        {
            // Count all OfficeMath nodes that belong to this section.
            int equationCount = section.Body.GetChildNodes(NodeType.OfficeMath, true).Count;

            summaryBuilder.AppendLine($"Section {sectionNumber}: {equationCount} equation(s)");
            sectionNumber++;
        }

        // Create a new blank document to hold the summary.
        Document summaryDoc = new Document();

        // Use DocumentBuilder to write the summary text into the document.
        DocumentBuilder builder = new DocumentBuilder(summaryDoc);
        builder.Writeln("OfficeMath Equation Summary");
        builder.Writeln(); // blank line
        builder.Writeln(summaryBuilder.ToString());

        // Save the summary document.
        summaryDoc.Save("OfficeMathSummary.docx");
    }
}
