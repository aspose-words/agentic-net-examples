using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathToggleExample
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with some text.
        builder.Writeln("Sample document with equations:");

        // Insert first equation (inline by default after toggling).
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction.

        // Insert a second equation.
        InsertEquation(builder, @"\r(3,x)"); // Simple root.

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath, SaveFormat.Docx);

        // Toggle display mode of top‑level OfficeMath equations.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = om.DisplayType == OfficeMathDisplayType.Inline
                    ? OfficeMathDisplayType.Display
                    : OfficeMathDisplayType.Inline;
            }
        }

        // Save the document after toggling.
        string toggledPath = "Toggled.docx";
        doc.Save(toggledPath, SaveFormat.Docx);

        // Simple validation that files were created.
        if (!File.Exists(originalPath) || !File.Exists(toggledPath))
        {
            throw new InvalidOperationException("One or more output files were not created.");
        }
    }

    // Helper method that creates a real OfficeMath node using the EQ‑field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is FieldEQ fieldEq)
        {
            // Write the EQ argument into the field separator.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqArgument);

            // Convert the field to an OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                Node fieldStart = fieldEq.Start;
                Node parent = fieldStart.ParentNode;

                if (parent is CompositeNode compositeParent)
                {
                    compositeParent.InsertBefore(officeMath, fieldStart);
                }
                else
                {
                    throw new InvalidOperationException("Parent node is not a composite node.");
                }

                // Remove the original field.
                fieldEq.Remove();
            }
        }
    }
}
