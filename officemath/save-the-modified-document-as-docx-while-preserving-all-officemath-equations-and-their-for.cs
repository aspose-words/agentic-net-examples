using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Output file path.
        string outputPath = "ModifiedDocument.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two equations using the deterministic EQ‑field bootstrap workflow.
        InsertOfficeMath(builder, @"\f(1,2)"); // Fraction 1/2
        InsertOfficeMath(builder, @"\r(3,x)"); // Cube root of x

        // Save the document as DOCX. All OfficeMath nodes are preserved with their formatting.
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    // Inserts an EQ field, converts it to a real OfficeMath object,
    // applies display formatting, and removes the original field.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments (the equation) into the field separator.
        if (field.Separator != null)
        {
            builder.MoveTo(field.Separator);
            builder.Write(eqArguments);
        }

        // Move back to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a paragraph break after the equation for readability.
        builder.InsertParagraph();

        // Update the field so that Word processes the EQ code.
        field.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Ensure conversion succeeded.
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);

        // Remove the original EQ field from the document.
        field.Remove();

        // Apply display formatting to top‑level OfficeMath nodes only.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;
        }
    }
}
