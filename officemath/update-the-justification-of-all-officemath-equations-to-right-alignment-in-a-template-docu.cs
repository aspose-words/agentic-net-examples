using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Simple fraction 1/2
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x

        // Convert all inserted EQ fields to real OfficeMath objects and remove the fields.
        foreach (FieldEQ fieldEq in doc.Range.Fields.OfType<FieldEQ>().ToList())
        {
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
                // Remove the original field.
                fieldEq.Remove();
            }
        }

        // Update justification of all top‑level OfficeMath equations to right alignment.
        foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true).OfType<OfficeMath>())
        {
            // Target only top‑level equations (MathObjectType.OMathPara).
            if (officeMath.MathObjectType == MathObjectType.OMathPara)
            {
                // Ensure the equation is in display mode before setting justification.
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.Right;
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");

        // Optional validation: confirm that each top‑level OfficeMath has right justification.
        foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true).OfType<OfficeMath>())
        {
            if (officeMath.MathObjectType == MathObjectType.OMathPara &&
                officeMath.Justification != OfficeMathJustification.Right)
                throw new Exception("Justification update failed for an equation.");
        }
    }

    // Helper method to insert an EQ field with the specified arguments and convert it later.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return to the field start's parent node and start a new paragraph for the next equation.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
    }
}
