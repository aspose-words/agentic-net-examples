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

        // Insert a paragraph to hold the equation.
        builder.Writeln("Equation with centered justification:");

        // Insert an EQ field which will be converted to OfficeMath.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write a simple equation (fraction).
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that it can be converted to a real OfficeMath node.
        field.Update();

        // Cast the field to FieldEQ.
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to cast field to FieldEQ.");

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        Node fieldStart = field.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        field.Remove();

        // Set justification only on top‑level equations.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.Justification = OfficeMathJustification.Center;
        }

        // Save the document.
        string outputPath = "JustifiedEquation.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the document to verify the justification persisted.
        Document loadedDoc = new Document(outputPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes found in the saved document.");

        OfficeMath savedMath = mathNodes[0] as OfficeMath;
        if (savedMath == null || savedMath.Justification != OfficeMathJustification.Center)
            throw new InvalidOperationException("Justification was not applied correctly.");

        Console.WriteLine("Document saved and equation justification set to center.");
    }
}
