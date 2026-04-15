using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Simple fraction.
        InsertEquation(builder, @"\r(3,x)");          // Cube root.
        InsertEquation(builder, @"\a \co2 (1,2,3,4)"); // 2‑column array (matrix‑like).

        // Save the sample document (optional, just to visualize the result).
        string samplePath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathSample.docx");
        doc.Save(samplePath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        Console.WriteLine("Scanning OfficeMath nodes for unsupported MathObjectTypes...");

        foreach (OfficeMath officeMath in mathNodes)
        {
            // The MathObjectType indicates the kind of math object.
            MathObjectType type = officeMath.MathObjectType;

            // Treat OMathPara (top‑level equation) as the supported type.
            if (type != MathObjectType.OMathPara)
            {
                // Log the unsupported type for review.
                Console.WriteLine($"Unsupported MathObjectType detected: {type}");
            }
        }

        // Indicate completion.
        Console.WriteLine("Processing completed.");
    }

    // Helper that inserts an EQ field, writes the argument string,
    // converts it to a real OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the field start position.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        }

        // Remove the original field from the document.
        field.Remove();

        // Add a paragraph break after the equation for readability.
        builder.Writeln();
    }
}
