using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class RetrieveMathObjectTypes
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a fraction equation: \f(1,2)
        InsertEquation(builder, @"\f(1,2)");

        // Insert a radical equation: \r(3,x)
        InsertEquation(builder, @"\r(3,x)");

        // Save the document so we can verify that the equations were created.
        const string outputPath = "MathObjects.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to save the document.");

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int index = 1;
        foreach (OfficeMath om in mathNodes)
        {
            // Determine the type of the OfficeMath node.
            MathObjectType type = om.MathObjectType;
            string description = type switch
            {
                MathObjectType.Fraction => "Fraction",
                MathObjectType.Radical => "Radical",
                _ => $"Other ({type})"
            };

            Console.WriteLine($"OfficeMath node #{index}: {description}");
            index++;
        }
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that the EQ code is parsed.
        field.Update();

        // Return the builder to the paragraph after the field.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();
    }
}
