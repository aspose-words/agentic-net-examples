using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    // Returns true if the given OfficeMath node has the specified MathObjectType.
    public static bool IsOfficeMathOfType(OfficeMath officeMath, MathObjectType expectedType)
    {
        if (officeMath == null)
            throw new ArgumentNullException(nameof(officeMath));

        return officeMath.MathObjectType == expectedType;
    }

    // Inserts an EQ field with the provided arguments, converts it to a real OfficeMath node,
    // replaces the field with the OfficeMath node, and returns the created OfficeMath.
    private static OfficeMath InsertOfficeMath(Document doc, string eqArguments)
    {
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that the EQ code is recognized.
        field.Update();

        // Return the builder to the field start's parent paragraph.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a new paragraph after the field (optional, just to keep layout clean).
        builder.InsertParagraph();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        return officeMath;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a simple fraction equation: \f(1,2)
        OfficeMath mathNode = InsertOfficeMath(doc, @"\f(1,2)");

        // Check if the inserted OfficeMath node is a top‑level paragraph equation (OMathPara).
        bool isPara = IsOfficeMathOfType(mathNode, MathObjectType.OMathPara);
        Console.WriteLine($"OfficeMath is of type OMathPara: {isPara}");

        // Save the document to verify the result.
        string outputPath = "OfficeMathSample.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
