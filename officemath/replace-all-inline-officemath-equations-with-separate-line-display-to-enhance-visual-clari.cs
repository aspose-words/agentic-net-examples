using System;
using System.IO;
using System.Linq;
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

        // First paragraph with an inline equation.
        builder.Writeln("Paragraph with an inline equation:");
        builder.Write("The fraction ");
        InsertInlineEquation(builder, @"\f(1,2)"); // creates 1/2 as inline equation
        builder.Writeln(" appears here.");

        // Second paragraph with another inline equation.
        builder.Writeln("Another paragraph:");
        builder.Write("Integral example: ");
        InsertInlineEquation(builder, @"\i \su(n=1,5,n)"); // creates a summation integral as inline equation
        builder.Writeln(" end of line.");

        // Process the document: replace all inline OfficeMath equations with display equations.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true).OfType<OfficeMath>())
        {
            // Target only top‑level equations (OMathPara) that are currently inline.
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType == OfficeMathDisplayType.Inline)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Output.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }

    // Helper method that inserts an EQ field, converts it to a real OfficeMath node,
    // sets it to inline display, and returns the created OfficeMath object.
    private static OfficeMath InsertInlineEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ arguments (e.g., "\f(1,2)").
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return the builder to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();

            // Ensure the equation is initially inline.
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }

        return officeMath;
    }
}
