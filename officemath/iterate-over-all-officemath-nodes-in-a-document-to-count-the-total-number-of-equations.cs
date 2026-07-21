using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathCounter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");   // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the document (optional, demonstrates the save rule).
        const string outputPath = "OfficeMathCount.docx";
        doc.Save(outputPath);

        // Count top‑level OfficeMath nodes (MathObjectType == OMathPara) – these represent equations.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the result.
        Console.WriteLine($"Total number of equations: {equationCount}");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArgs)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgs);

        // Return the cursor to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }
    }
}
