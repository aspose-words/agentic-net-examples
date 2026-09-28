using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    // Returns true if the OfficeMath node's MathObjectType matches the specified type.
    public static bool IsMathObjectType(OfficeMath officeMath, MathObjectType type)
    {
        if (officeMath == null)
            throw new ArgumentNullException(nameof(officeMath));

        return officeMath.MathObjectType == type;
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will be converted to a real OfficeMath object.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator to write the EQ argument.
        builder.MoveTo(field.Separator);
        // Write a simple, safe equation (fraction 1 over 2).
        builder.Write(@"\f(1,2)");

        // Update the field so that the equation is evaluated.
        field.Update();

        // Convert the field to OfficeMath.
        FieldEQ fieldEq = field as FieldEQ;
        OfficeMath officeMath = fieldEq?.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            CompositeNode parent = (CompositeNode)field.Start.ParentNode;
            parent.InsertBefore(officeMath, field.Start);

            // Remove the original field.
            field.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Save the document.
        const string outputPath = "Sample.docx";
        doc.Save(outputPath);

        // Reload the document to demonstrate loading.
        Document loadedDoc = new Document(outputPath);

        // Get all OfficeMath nodes in the document.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Check each OfficeMath node against the OMathPara type.
        foreach (OfficeMath om in mathNodes)
        {
            bool isPara = IsMathObjectType(om, MathObjectType.OMathPara);
            Console.WriteLine($"OfficeMath node MathObjectType is OMathPara: {isPara}");
        }
    }
}
