using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class SetOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that introduces the equation.
        builder.Writeln("Equation below:");

        // Insert a deterministic EQ field that will be converted to a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = field as FieldEQ;
        if (fieldEQ == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write a simple fraction equation into the field separator.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that Aspose.Words can convert it to OfficeMath.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        Node fieldStart = fieldEQ.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        fieldEQ.Remove();

        // Ensure we are dealing with a top‑level equation (OMathPara) and set its justification to center.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.Justification = OfficeMathJustification.Center;
        }
        else
        {
            throw new InvalidOperationException("Created OfficeMath is not a top‑level paragraph.");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Output file was not created.", outputPath);

        // Reload the document and confirm the justification.
        Document loadedDoc = new Document(outputPath);
        OfficeMath loadedMath = (OfficeMath)loadedDoc.GetChildNodes(NodeType.OfficeMath, true)[0];
        if (loadedMath.Justification != OfficeMathJustification.Center)
            throw new InvalidOperationException("Justification was not set to center.");

        // Example completed successfully.
    }
}
