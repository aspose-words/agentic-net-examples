using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ReplaceInlineOfficeMath
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text and an inline equation.
        builder.Writeln("Paragraph with first inline equation:");
        InsertInlineEquation(builder, @"\f(1,2)"); // Fraction 1/2

        // Insert another paragraph with text and a second inline equation.
        builder.Writeln("Paragraph with second inline equation:");
        InsertInlineEquation(builder, @"\r(3,x)"); // Cube root of x

        // Insert a third paragraph with text and a third inline equation.
        builder.Writeln("Paragraph with third inline equation:");
        InsertInlineEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the intermediate document (optional, for inspection).
        string intermediatePath = "Intermediate.docx";
        doc.Save(intermediatePath);

        // Iterate over all OfficeMath nodes and change inline display to display mode.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            // Target only top‑level equations (MathObjectType.OMathPara).
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // If the equation is currently inline, switch it to display mode.
                if (om.DisplayType == OfficeMathDisplayType.Inline)
                {
                    om.DisplayType = OfficeMathDisplayType.Display;
                    om.Justification = OfficeMathJustification.Left;
                }
            }
        }

        // Save the final document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    // Helper method to insert an EQ field, convert it to OfficeMath, and keep it inline.
    private static void InsertInlineEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return the builder to the field start's parent (the current paragraph).
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();

            // Ensure the equation is treated as inline initially.
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Continue building in the same paragraph.
        builder.Writeln(); // Add a line break after the equation to keep the flow readable.
    }
}
