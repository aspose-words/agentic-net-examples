using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Toggle the display mode of each top‑level OfficeMath node.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Switch between Inline and Display.
                if (om.DisplayType == OfficeMathDisplayType.Inline)
                {
                    om.DisplayType = OfficeMathDisplayType.Display;
                    om.Justification = OfficeMathJustification.Left;
                }
                else
                {
                    om.DisplayType = OfficeMathDisplayType.Inline;
                    // Justification cannot be set when Inline, so we leave it unchanged.
                }
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ToggledOfficeMath.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and returns the node.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that its internal code reflects the arguments.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Ensure conversion succeeded.
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        // Move the builder to the paragraph after the inserted equation for further insertions.
        builder.MoveTo(officeMath.ParentParagraph);
        builder.InsertParagraph();

        return officeMath;
    }
}
