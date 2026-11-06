using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field equation (EQ field) that will later be converted to a real OfficeMath node.
        FieldEQ field = builder.InsertField(FieldType.FieldEquation, true) as FieldEQ;

        // Move to the field separator and write a simple fraction equation.
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)"); // Fraction 1/2

        // Update the field so that the EQ switch is processed.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        Node fieldStart = field.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        field.Remove();

        // Apply center justification to top‑level equations only.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
            officeMath.Justification = OfficeMathJustification.Center;

        // Save the document.
        string outputPath = "OfficeMathJustified.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Output file was not created.");

        // Reload the document and confirm the justification.
        Document loadedDoc = new Document(outputPath);
        NodeCollection omaths = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        bool centeredFound = false;
        foreach (OfficeMath om in omaths)
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.Justification == OfficeMathJustification.Center)
            {
                centeredFound = true;
                break;
            }
        }

        if (!centeredFound)
            throw new InvalidOperationException("OfficeMath justification was not set to center.");
    }
}
