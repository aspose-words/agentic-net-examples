using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathMatcher
{
    // Returns true if the OfficeMath node's MathObjectType matches the specified criteria.
    public static bool MatchesMathObjectType(OfficeMath officeMath, MathObjectType criteria)
    {
        if (officeMath == null)
            return false;

        return officeMath.MathObjectType == criteria;
    }

    // Helper to create a simple OfficeMath equation using the deterministic EQ-field bootstrap workflow.
    private static OfficeMath CreateOfficeMath(Document doc, string eqArguments)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Ensure the field is up‑to‑date before conversion.
        field.Update();

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath before the field start.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);

        // Remove the original field.
        field.Remove();

        return officeMath;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a simple fraction equation: \f(1,2)
        OfficeMath mathNode = CreateOfficeMath(doc, @"\f(1,2)");

        // Save the document (optional, for verification).
        doc.Save("OfficeMathSample.docx");

        // Retrieve the first OfficeMath node from the document.
        OfficeMath retrievedMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

        // Check if it matches the OMathPara type.
        bool isOMathPara = MatchesMathObjectType(retrievedMath, MathObjectType.OMathPara);
        Console.WriteLine($"MathObjectType is OMathPara: {isOMathPara}");

        // Example of checking a different type (e.g., Fraction).
        bool isFraction = MatchesMathObjectType(retrievedMath, MathObjectType.Fraction);
        Console.WriteLine($"MathObjectType is Fraction: {isFraction}");
    }
}
