using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class ReplaceInlineOfficeMath
{
    public static void Main()
    {
        // Paths for the sample and result documents.
        const string samplePath = "Sample.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX with inline OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with an inline equation.
        builder.Writeln("Paragraph with an inline equation:");
        InsertInlineEquation(builder, @"\f(1,2)"); // Simple fraction.

        // Second paragraph with another inline equation.
        builder.Writeln("Another paragraph containing an inline equation:");
        InsertInlineEquation(builder, @"\r(3,x)"); // Simple radical.

        // Save the sample document.
        doc.Save(samplePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Reload the document and replace inline equations with display mode.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);

        // Find all top‑level OfficeMath nodes (MathObjectType == OMathPara) that are inline.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType == OfficeMathDisplayType.Inline)
            {
                // Change the display type to separate line (Display).
                om.DisplayType = OfficeMathDisplayType.Display;
            }
        }

        // Save the modified document.
        loadedDoc.Save(resultPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 3: Validation.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new Exception("Result document was not created.");

        Document verifyDoc = new Document(resultPath);
        NodeCollection verifyMath = verifyDoc.GetChildNodes(NodeType.OfficeMath, true);
        bool hasDisplay = false;
        foreach (OfficeMath om in verifyMath)
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType == OfficeMathDisplayType.Display)
            {
                hasDisplay = true;
                break;
            }
        }

        if (!hasDisplay)
            throw new Exception("No OfficeMath equation was set to display mode.");

        // Execution completed successfully.
    }

    // Helper method to insert an inline OfficeMath equation using the EQ‑field bootstrap workflow.
    private static void InsertInlineEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an empty Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = (FieldEQ)field;

        // Write the EQ argument into the field separator.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(eqArgument);

        // Update the field so that Aspose.Words can convert it to a real OfficeMath node.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new Exception("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        CompositeNode parent = (CompositeNode)fieldEq.Start.ParentNode;
        parent.InsertBefore(officeMath, fieldEq.Start);
        fieldEq.Remove();
    }
}
