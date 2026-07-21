using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        const string filePath = "OfficeMathTypes.docx";

        // Create a new document and insert sample equations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Fraction: 1/2
        InsertEquation(builder, @"\f(1,2)");

        // Radical: cube root of x
        InsertEquation(builder, @"\r(3,x)");

        // Save the document.
        doc.Save(filePath);

        // Reload the document to demonstrate loading.
        Document loadedDoc = new Document(filePath);

        // Retrieve all OfficeMath nodes.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Output the MathObjectType of each node.
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            string description = om.MathObjectType switch
            {
                MathObjectType.Fraction => "Fraction",
                MathObjectType.Radical => "Radical",
                _ => om.MathObjectType.ToString()
            };
            Console.WriteLine($"Node {i}: {description}");
        }
    }

    // Helper that inserts an EQ field, converts it to OfficeMath, and moves to a new paragraph.
    private static void InsertEquation(DocumentBuilder builder, string eqArgs)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgs);

        // Return to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath before the field and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Start a new paragraph for the next equation.
        builder.InsertParagraph();
    }
}
