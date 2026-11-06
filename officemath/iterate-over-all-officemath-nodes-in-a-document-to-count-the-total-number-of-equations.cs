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

        // Add some text and equations using the deterministic EQ‑field bootstrap workflow.
        builder.Writeln("Paragraph before the first equation.");
        InsertEquation(builder, @"\f(1,2)"); // Simple fraction.

        builder.Writeln("Paragraph between equations.");
        InsertEquation(builder, @"\r(3,x)"); // Cube root.

        builder.Writeln("Paragraph after the second equation.");

        // Save the document.
        string outputPath = "SampleWithEquations.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved.");

        // Count top‑level equations (OfficeMath nodes whose MathObjectType is OMathPara).
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the result.
        Console.WriteLine($"Total equations: {equationCount}");
    }

    // Helper that creates a real OfficeMath node from an EQ field and inserts it into the document.
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an empty EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move the builder to the field separator node and write the EQ argument.
        // The separator node is where the field's code (the EQ switch) should be placed.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);

        // Convert the field to an OfficeMath node.
        FieldEQ fieldEq = field as FieldEQ;
        OfficeMath officeMath = fieldEq?.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            Node fieldStart = field.Start;
            CompositeNode parent = fieldStart.ParentNode as CompositeNode;
            if (parent != null)
                parent.InsertBefore(officeMath, fieldStart);

            // Remove the original field, leaving only the real equation.
            field.Remove();
        }
    }
}
