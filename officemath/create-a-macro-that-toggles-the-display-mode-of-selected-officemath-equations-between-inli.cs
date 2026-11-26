using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ToggleOfficeMathDisplay
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some introductory text.
        builder.Writeln("Sample equations:");

        // Insert a few equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Toggle the display mode of each top‑level OfficeMath equation.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            // Operate only on top‑level equations (MathObjectType.OMathPara).
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                if (om.DisplayType == OfficeMathDisplayType.Inline)
                {
                    // Change to display (separate line) mode.
                    om.DisplayType = OfficeMathDisplayType.Display;
                    om.Justification = OfficeMathJustification.Left;
                }
                else
                {
                    // Change to inline mode.
                    om.DisplayType = OfficeMathDisplayType.Inline;
                    om.Justification = OfficeMathJustification.Inline;
                }
            }
        }

        // Save the resulting document.
        const string outputPath = "ToggledOfficeMath.docx";
        doc.Save(outputPath);
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eq)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ argument string into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eq);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // If conversion succeeded, replace the field with the OfficeMath node.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a line break after the equation for readability.
        builder.Writeln();
    }
}
