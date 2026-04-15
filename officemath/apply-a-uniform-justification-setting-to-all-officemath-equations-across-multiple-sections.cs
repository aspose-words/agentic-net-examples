using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ApplyOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample document with three sections, each containing two equations.
        for (int sectionIndex = 0; sectionIndex < 3; sectionIndex++)
        {
            if (sectionIndex > 0)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln($"Section {sectionIndex + 1} – introductory text.");

            // First equation in the section.
            InsertEquation(builder, @"\f(1,2)"); // Simple fraction 1/2.
            builder.Writeln(); // Move to next line.

            // Second equation in the section.
            InsertEquation(builder, @"\r(3,x)"); // Cube root of x.
            builder.Writeln(); // End of section content.
        }

        // Apply a uniform justification (Center) to all top‑level OfficeMath paragraphs.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes.OfType<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Ensure the equation is displayed on its own line before setting justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Center;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UniformJustifiedEquations.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation: confirm the file exists and each top‑level equation has the expected justification.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        Document loaded = new Document(outputPath);
        NodeCollection loadedMath = loaded.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in loadedMath.OfType<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                if (om.Justification != OfficeMathJustification.Center)
                    throw new InvalidOperationException("An equation does not have the expected justification.");
            }
        }
    }

    // Helper that creates a real OfficeMath node using the deterministic EQ‑field bootstrap workflow.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments (e.g., "\f(1,2)").
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field to ensure the code is processed.
        field.Update();

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the OfficeMath node.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        return officeMath;
    }
}
