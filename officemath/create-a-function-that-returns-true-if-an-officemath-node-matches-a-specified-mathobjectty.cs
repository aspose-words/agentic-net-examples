using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that represents a simple fraction: \f(1,2)
        FieldEQ field = InsertFieldEQ(builder, @"\f(1,2)");

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the original field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Save the document to a local folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        doc.Save(docPath);

        // Retrieve the first OfficeMath node from the document.
        OfficeMath firstMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

        // Check whether the node's MathObjectType matches OMathPara.
        bool matches = IsMathObjectType(firstMath, MathObjectType.OMathPara);
        Console.WriteLine($"OfficeMath matches OMathPara: {matches}");
    }

    // Helper method to insert an EQ field with the specified arguments.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(field.Separator);
        builder.Write(args);
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
        return field;
    }

    // Returns true if the given OfficeMath node's MathObjectType equals the target type.
    private static bool IsMathObjectType(OfficeMath officeMath, MathObjectType targetType)
    {
        if (officeMath == null)
            return false;

        return officeMath.MathObjectType == targetType;
    }
}
