using System;
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

        // Insert a fraction equation: 1/2
        InsertOfficeMath(builder, @"\f(1,2)");
        builder.Writeln(); // separate equations with a new paragraph

        // Insert a radical equation: cube root of x
        InsertOfficeMath(builder, @"\r(3,x)");

        // Save the document (optional, just to demonstrate saving).
        const string outputPath = "OfficeMathTypes.docx";
        doc.Save(outputPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int index = 0;
        foreach (OfficeMath officeMath in mathNodes)
        {
            MathObjectType type = officeMath.MathObjectType;
            string description = type == MathObjectType.Fraction ? "Fraction"
                                 : type == MathObjectType.Radical ? "Radical"
                                 : type.ToString();

            Console.WriteLine($"OfficeMath node {index}: {description}");
            index++;
        }
    }

    // Helper that creates a real OfficeMath object from an EQ field using the deterministic bootstrap workflow.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ switch arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the field start location.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Replace the field with the generated OfficeMath node.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }
    }
}
