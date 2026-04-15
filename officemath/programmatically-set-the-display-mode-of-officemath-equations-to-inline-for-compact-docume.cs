using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class SetOfficeMathDisplayInline
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some introductory text.
        builder.Writeln("Sample document with equations:");

        // Insert a few equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)"); // Fraction 1/2
        builder.Writeln(); // New line after the equation.

        InsertEquation(builder, @"\r(3,x)"); // Cube root of x
        builder.Writeln();

        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation.
        builder.Writeln();

        // Change the display mode of all top‑level OfficeMath nodes to Inline.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Inline;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // The program finishes without waiting for user input.
    }

    // Helper that creates a real OfficeMath node from an EQ field and returns it.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the arguments for the EQ field.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        return officeMath;
    }
}
