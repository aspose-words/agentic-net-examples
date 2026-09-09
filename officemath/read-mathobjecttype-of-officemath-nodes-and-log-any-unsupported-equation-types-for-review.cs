using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        // 1. Simple fraction.
        InsertEquation(builder, @"\f(1,2)");
        // 2. Integral with summation.
        InsertEquation(builder, @"\i \su(n=1,5,n)");
        // 3. Matrix (array of equations).
        InsertEquation(builder, @"\a \co2 \vs1 \hs1(1,2,3,4)");
        // 4. Radical.
        InsertEquation(builder, @"\r(3,x)");
        // 5. Function.
        InsertEquation(builder, @"\f(x) = \f(x^2)");

        // Save the document (optional, just to have a file to inspect).
        string docPath = Path.Combine(artifactsDir, "OfficeMathSample.docx");
        doc.Save(docPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        Console.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");

        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath officeMath = (OfficeMath)mathNodes[i];
            MathObjectType type = officeMath.MathObjectType;

            // We consider OMathPara (display equations) as supported.
            if (type != MathObjectType.OMathPara)
            {
                Console.WriteLine($"Unsupported MathObjectType at index {i}: {type}");
            }
            else
            {
                Console.WriteLine($"Supported MathObjectType at index {i}: {type}");
            }
        }
    }

    // Helper that inserts an EQ field, converts it to OfficeMath, and cleans up the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Move back to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
