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
        // Paths for the sample and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Updated.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document with a few OfficeMath equations.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three simple equations, each in its own paragraph.
        InsertEquation(builder, @"\f(1,2)"); // fraction 1/2
        InsertEquation(builder, @"\r(3,x)"); // cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // integral with summation

        // Save the template document.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Reload the document and update justification of all top‑level equations.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(templatePath);

        // Find all OfficeMath nodes.
        var officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true)
                                      .Cast<OfficeMath>()
                                      .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in officeMathNodes)
        {
            // Ensure the equation is displayed on its own line before setting justification.
            om.DisplayType = OfficeMathDisplayType.Display;
            om.Justification = OfficeMathJustification.Right;
        }

        // Save the modified document.
        loadedDoc.Save(outputPath);

        // -----------------------------------------------------------------
        // 3. Validate that the justification was applied.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Output file '{outputPath}' was not created.");

        Document resultDoc = new Document(outputPath);
        var resultMaths = resultDoc.GetChildNodes(NodeType.OfficeMath, true)
                                   .Cast<OfficeMath>()
                                   .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in resultMaths)
        {
            if (om.Justification != OfficeMathJustification.Right)
                throw new InvalidOperationException("One or more equations do not have right justification.");
        }

        // All done.
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and moves the builder to the next paragraph.
    private static void InsertEquation(DocumentBuilder builder, string eqArgs)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ argument string.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgs);
        // Return to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);
        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }
        // Start a new paragraph for the next equation.
        builder.InsertParagraph();
    }
}
