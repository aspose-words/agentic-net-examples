using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class UpdateOfficeMathJustification
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "UpdatedOfficeMath.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)"); // Cube root of x
        InsertEquation(builder, @"\i");      // Integral symbol

        // Traverse all OfficeMath nodes in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in officeMathNodes)
        {
            // Apply changes only to top‑level equations (MathObjectType == OMathPara).
            if (officeMath.MathObjectType == MathObjectType.OMathPara)
            {
                // Set display type before changing justification.
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.Right;
            }
        }

        // Save the modified document.
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArgs)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgs);

        // Return the builder to the field start position.
        builder.MoveTo(field.Start);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Move to a new paragraph for the next equation.
        builder.Writeln();
    }
}
