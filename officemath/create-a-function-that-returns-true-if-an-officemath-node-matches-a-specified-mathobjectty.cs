using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    // Returns true if the given OfficeMath node has the specified MathObjectType.
    public static bool IsMatchingMathObjectType(OfficeMath officeMath, MathObjectType targetType)
    {
        if (officeMath == null)
            return false;

        return officeMath.MathObjectType == targetType;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // The equation "\f(1,2)" creates a simple fraction 1/2.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ switch and its arguments at the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Return the builder to the paragraph that contains the field and add a new paragraph.
        builder.MoveTo(eqField.Start.ParentNode);
        builder.InsertParagraph();

        // Update the field so that its internal code is recognized.
        eqField.Update();

        // Convert the EQ field to OfficeMath.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Ensure conversion succeeded before proceeding.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the original field and remove the field.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
            eqField.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Save the document (optional, demonstrates that the output exists).
        const string outputPath = "OfficeMathSample.docx";
        doc.Save(outputPath);

        // Retrieve the first OfficeMath node from the document.
        OfficeMath firstOfficeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

        // Check if the node is a top‑level equation (MathObjectType.OMathPara).
        bool isTopLevel = IsMatchingMathObjectType(firstOfficeMath, MathObjectType.OMathPara);
        Console.WriteLine($"OfficeMath node is top‑level equation: {isTopLevel}");

        // Example of checking for a different type (e.g., Fraction).
        bool isFraction = IsMatchingMathObjectType(firstOfficeMath, MathObjectType.Fraction);
        Console.WriteLine($"OfficeMath node is a fraction: {isFraction}");
    }
}
