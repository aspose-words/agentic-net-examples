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

        // Build a sample document with three sections, each containing an equation.
        for (int sectionIndex = 1; sectionIndex <= 3; sectionIndex++)
        {
            // Add a heading for the section.
            builder.Writeln($"Section {sectionIndex}");

            // Insert a simple fraction equation using the deterministic EQ-field bootstrap workflow.
            InsertEquation(builder, @"\f(1,2)");

            // Add some regular text after the equation.
            builder.Writeln("Some explanatory text.");

            // Insert a section break after each section except the last one.
            if (sectionIndex < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Apply a uniform justification to all top‑level OfficeMath equations.
        const OfficeMathJustification targetJustification = OfficeMathJustification.Center;
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Display type must be set before justification when using Display mode.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = targetJustification;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Verify that every top‑level equation has the expected justification.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara && om.Justification != targetJustification)
                throw new InvalidOperationException("An equation does not have the expected justification.");
        }
    }

    // Inserts an EQ field with the given argument string, converts it to OfficeMath,
    // replaces the field with the real OfficeMath node, and returns the created OfficeMath.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Ensure the field is up‑to‑date before conversion.
        field.Update();

        // Return the builder to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        return officeMath;
    }
}
