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

        // Insert a fraction equation: 1/2
        InsertOfficeMath(builder, @"\f(1,2)");

        // Insert a radical equation: cube root of x
        InsertOfficeMath(builder, @"\r(3,x)");

        // Save the document to disk.
        string outputPath = "OfficeMathTypes.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            // Retrieve the MathObjectType of the node.
            MathObjectType type = om.MathObjectType;
            Console.WriteLine($"OfficeMath node: MathObjectType = {type}");

            // Determine whether the node is a fraction or a radical.
            if (type == MathObjectType.Fraction)
                Console.WriteLine("-> This node represents a fraction.");
            else if (type == MathObjectType.Radical)
                Console.WriteLine("-> This node represents a radical.");
            else
                Console.WriteLine("-> This node is of another type.");
        }
    }

    // Helper method that creates a real OfficeMath node from an EQ field using the deterministic bootstrap workflow.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field placeholder.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments (e.g., "\f(1,2)" or "\r(3,x)").
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // If conversion succeeded, replace the field with the OfficeMath node.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
