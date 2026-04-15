using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Drawing;

public class OfficeMathCounter
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Simple fraction 1/2
        builder.Writeln();                         // New paragraph
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        builder.Writeln();                         // New paragraph
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Count top‑level OfficeMath nodes (MathObjectType.OMathPara) – these represent equations.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the count.
        Console.WriteLine($"Total number of equations: {equationCount}");

        // Save the document so the example produces an output file.
        doc.Save("SampleWithEquations.docx");
    }

    // Helper that creates a real OfficeMath node from an EQ field and inserts it into the document.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        }

        // Remove the original field from the document.
        field.Remove();

        return officeMath;
    }
}
