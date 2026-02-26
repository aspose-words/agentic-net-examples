using System;
using Aspose.Words;
using Aspose.Words.Math;

class OfficeMathSummary
{
    static void Main()
    {
        // Load the DOT template document.
        Document template = new Document("Template.dotx");

        // Create a new blank document that will hold the summary.
        Document summary = new Document();
        DocumentBuilder builder = new DocumentBuilder(summary);

        // Header for the summary.
        builder.Writeln("OfficeMath Equation Count per Section");
        builder.Writeln("-------------------------------------");

        // Iterate through each section of the template.
        for (int i = 0; i < template.Sections.Count; i++)
        {
            // Current section.
            var section = template.Sections[i];

            // Count OfficeMath objects inside this section (including nested ones).
            int mathCount = section.Body.GetChildNodes(NodeType.OfficeMath, true).Count;

            // Write the count to the summary document.
            builder.Writeln($"Section {i + 1}: {mathCount} equation(s)");
        }

        // Save the summary document.
        summary.Save("OfficeMathSummary.docx");
    }
}
