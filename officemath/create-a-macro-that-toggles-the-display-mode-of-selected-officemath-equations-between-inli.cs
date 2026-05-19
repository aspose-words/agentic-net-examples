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
        // The equation "\f(1,2)" creates a simple fraction 1/2.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");
        builder.MoveTo(eqField.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
            eqField.Remove();
        }

        // Add another equation to demonstrate toggling multiple equations.
        builder.Writeln("Another equation:");
        FieldEQ eqField2 = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(eqField2.Separator);
        builder.Write(@"\r(3,x)"); // Cube root of x.
        builder.MoveTo(eqField2.Start.ParentNode);
        builder.InsertParagraph();

        OfficeMath officeMath2 = eqField2.AsOfficeMath();
        if (officeMath2 != null)
        {
            eqField2.Start.ParentNode.InsertBefore(officeMath2, eqField2.Start);
            eqField2.Remove();
        }

        // Toggle display mode for all top‑level OfficeMath equations (MathObjectType.OMathPara).
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
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
                    om.Justification = OfficeMathJustification.Inline;
                }
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathToggled.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
