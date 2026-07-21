using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    // Returns true if the given OfficeMath node has the specified MathObjectType.
    public static bool IsMathObjectType(OfficeMath officeMath, MathObjectType targetType)
    {
        if (officeMath == null)
            return false;

        return officeMath.MathObjectType == targetType;
    }

    // Helper that creates a real OfficeMath node using the deterministic EQ‑field bootstrap workflow.
    private static OfficeMath InsertFractionEquation(DocumentBuilder builder, string fractionArgs)
    {
        // Insert an empty EQ field.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ switch arguments (e.g. "\f(1,2)").
        builder.MoveTo(eqField.Separator);
        builder.Write(fractionArgs);

        // Update the field so that its internal state is consistent.
        eqField.Update();

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);
        builder.InsertParagraph(); // Optional separation.

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the real OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        return officeMath;
    }

    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple fraction equation: \f(1,2)
        OfficeMath firstMath = InsertFractionEquation(builder, @"\f(1,2)");

        // Verify the type of the created OfficeMath node.
        bool isPara = IsMathObjectType(firstMath, MathObjectType.OMathPara);
        Console.WriteLine($"OfficeMath node is OMathPara: {isPara}");

        bool isFraction = IsMathObjectType(firstMath, MathObjectType.Fraction);
        Console.WriteLine($"OfficeMath node is Fraction: {isFraction}");

        // Save the document to demonstrate persistence.
        doc.Save("OfficeMathExample.docx");
    }
}
