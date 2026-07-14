using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class InsertOfficeMathFromMathML
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the equation.
        builder.Writeln("Paragraph before the equation.");

        // Insert an empty paragraph where the equation will be placed.
        builder.Writeln();

        // Insert an EQ field – the deterministic way to bootstrap a real OfficeMath node.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write a simple EQ argument at the field separator.
        // This argument will be converted to a real OfficeMath object.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // simple fraction 1/2

        // Update the field so that the EQ code is processed.
        eqField.Update();

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Replace the field with the generated OfficeMath node.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
            // Remove the original field from the document.
            eqField.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Additional validation: ensure the document now contains a top‑level OfficeMath node.
        OfficeMath foundMath = doc.GetChild(NodeType.OfficeMath, 0, true) as OfficeMath;
        if (foundMath == null || foundMath.MathObjectType != MathObjectType.OMathPara)
            throw new InvalidOperationException("The expected OfficeMath node was not found in the document.");
    }
}
