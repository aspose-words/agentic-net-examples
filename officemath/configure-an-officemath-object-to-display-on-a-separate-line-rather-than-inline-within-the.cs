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

        // Add a paragraph with introductory text.
        builder.Writeln("Below is a simple equation displayed on its own line:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // Use a safe fraction equation: \f(1,2)
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Ensure the field is up‑to‑date before conversion.
        eqField.Update();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Only top‑level OfficeMath (OMathPara) can have its DisplayType changed.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Validate that the OfficeMath node has the correct display type.
        OfficeMath savedMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
        if (savedMath == null || savedMath.DisplayType != OfficeMathDisplayType.Display)
            throw new InvalidOperationException("OfficeMath display type was not set to Display.");
    }
}
