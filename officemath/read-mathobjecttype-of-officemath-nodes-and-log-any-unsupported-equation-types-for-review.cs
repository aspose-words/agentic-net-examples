using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class OfficeMathTypeLogger
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");                     // Fraction
        InsertEquation(builder, @"\r(3,x)");                     // Radical
        InsertEquation(builder, @"\a \al \co2 \vs3 \hs3(4x,-4y,-4x,+y)"); // Array (matrix‑like)
        InsertEquation(builder, @"\i \su(n=1,5,n)");            // Integral with summation
        InsertEquation(builder, @"\s \up8(Superscript) \s \do8(Subscript)"); // Superscript & subscript

        // Save the document (optional, demonstrates that the file was created).
        const string outputPath = "OfficeMathSample.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        // Define which MathObjectTypes are considered supported.
        var supportedTypes = new HashSet<MathObjectType>
        {
            MathObjectType.OMathPara   // Top‑level equation paragraph.
        };

        // Log any unsupported MathObjectTypes.
        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            MathObjectType type = om.MathObjectType;

            if (!supportedTypes.Contains(type))
            {
                Console.WriteLine($"Unsupported MathObjectType: {type} (Node index {i})");
            }
        }
    }

    // Helper that inserts an EQ field, writes its arguments, converts it to OfficeMath,
    // inserts the resulting OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // If conversion succeeded, replace the field with the OfficeMath node.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field from the document.
            field.Remove();
        }

        // Move the builder to the end of the paragraph to start a new one for the next equation.
        builder.MoveTo(officeMath?.ParentParagraph ?? field.Start.ParentNode);
        builder.Writeln(); // Ensure the next equation starts on a new line.
    }
}
