using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathToggleExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("Sample paragraph with an equation:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // The equation is a simple fraction 1/2.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(1,2)");

        // Convert the EQ field to OfficeMath and replace the field with the OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
            // Remove the original field from the document.
            eqField.Remove();
        }

        // Add another paragraph with a second equation (integral).
        builder.Writeln();
        builder.Writeln("Another equation:");
        FieldEQ eqField2 = InsertFieldEQ(builder, @"\i \su(n=1,5,n)");
        OfficeMath officeMath2 = eqField2.AsOfficeMath();
        if (officeMath2 != null)
        {
            eqField2.Start.ParentNode.InsertBefore(officeMath2, eqField2.Start);
            eqField2.Remove();
        }

        // Toggle the display mode of all top‑level OfficeMath equations.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            // Only modify top‑level equations (MathObjectType.OMathPara).
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                if (om.DisplayType == OfficeMathDisplayType.Inline)
                {
                    // Switch to display (separate line) mode.
                    om.DisplayType = OfficeMathDisplayType.Display;
                    // When in display mode a justification can be set.
                    om.Justification = OfficeMathJustification.Left;
                }
                else
                {
                    // Switch back to inline mode.
                    om.DisplayType = OfficeMathDisplayType.Inline;
                    // Do not set justification for inline mode (default is Inline).
                }
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ToggledOfficeMath.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method that inserts an EQ field with the specified arguments.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return the builder to the paragraph after the field.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a new paragraph after the equation for readability.
        builder.InsertParagraph();
        return field;
    }
}
