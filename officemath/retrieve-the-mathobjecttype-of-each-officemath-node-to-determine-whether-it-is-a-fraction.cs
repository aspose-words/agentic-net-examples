using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class OfficeMathTypeDemo
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

        // Save the document with the created equations.
        const string outputPath = "OfficeMathTypes.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Reload the document to demonstrate enumeration of OfficeMath nodes.
        Document loadedDoc = new Document(outputPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        Console.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            string typeDescription = om.MathObjectType switch
            {
                MathObjectType.Fraction => "Fraction",
                MathObjectType.Radical => "Radical",
                _ => $"Other ({om.MathObjectType})"
            };

            Console.WriteLine($"OfficeMath #{i + 1}: {typeDescription}");
        }
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node,
    // and removes the original field.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath before the field start node.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field from the document.
            field.Remove();
        }

        // Add a new paragraph after the inserted equation for readability.
        builder.InsertParagraph();
    }
}
