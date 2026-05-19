using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class UpdateOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several sample equations using the EQ field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction 1/2
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation
        InsertEquation(builder, @"\r(3,x)"); // Cube root of x

        // Convert all EQ fields to real OfficeMath objects and remove the original fields.
        foreach (FieldEQ fieldEq in doc.Range.Fields.OfType<FieldEQ>().ToList())
        {
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
                // Remove the original EQ field.
                fieldEq.Remove();
            }
        }

        // Update justification of all top‑level OfficeMath equations to right alignment.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Ensure the equation is in display mode before setting justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Right;
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "UpdatedOfficeMath.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Verify that each top‑level OfficeMath has the right justification.
        Document resultDoc = new Document(outputPath);
        foreach (OfficeMath om in resultDoc.GetChildNodes(NodeType.OfficeMath, true).Cast<OfficeMath>())
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                if (om.Justification != OfficeMathJustification.Right)
                    throw new Exception("Justification update failed for an equation.");
            }
        }

        // The program finishes without requiring any user interaction.
    }

    // Helper method to insert an EQ field with the specified arguments and convert it to a paragraph.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return to the field start's parent node.
        builder.MoveTo(field.Start.ParentNode);
        // End the paragraph after the equation.
        builder.InsertParagraph();
    }
}
