using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class Program
{
    // Returns true if the given OfficeMath node has the specified MathObjectType.
    public static bool IsMathObjectTypeMatch(OfficeMath officeMath, MathObjectType expectedType)
    {
        // Guard against null references.
        if (officeMath == null)
            return false;

        return officeMath.MathObjectType == expectedType;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // The equation is a simple fraction: 1 over 2.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");
        // Move back to the paragraph after the field to continue building if needed.
        builder.MoveTo(eqField.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the EQ field to an OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();

        // Ensure the conversion succeeded before inserting.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
            // Remove the original EQ field from the document.
            eqField.Remove();
        }

        // Save the document to a temporary file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathSample.docx");
        doc.Save(outputPath);

        // Reload the document to demonstrate discovery of the OfficeMath node.
        Document loadedDoc = new Document(outputPath);

        // Retrieve the first OfficeMath node in the document.
        OfficeMath firstOfficeMath = loadedDoc.GetChild(NodeType.OfficeMath, 0, true) as OfficeMath;

        // Check if the node is a top‑level equation (MathObjectType.OMathPara).
        bool isPara = IsMathObjectTypeMatch(firstOfficeMath, MathObjectType.OMathPara);
        Console.WriteLine($"OfficeMath node is OMathPara: {isPara}");

        // Example of checking for a different type (e.g., Fraction).
        bool isFraction = IsMathObjectTypeMatch(firstOfficeMath, MathObjectType.Fraction);
        Console.WriteLine($"OfficeMath node is Fraction: {isFraction}");
    }
}
