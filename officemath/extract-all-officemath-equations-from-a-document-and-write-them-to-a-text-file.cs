using System;
using System.Collections.Generic;
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

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertOfficeMath(builder, @"\f(1,2)");          // Fraction 1/2
        InsertOfficeMath(builder, @"\r(3,x)");          // Cube root of x
        InsertOfficeMath(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Optional: save the sample document (not required for extraction but demonstrates the workflow).
        string samplePath = "SampleWithEquations.docx";
        doc.Save(samplePath);

        // Extract all OfficeMath nodes from the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        List<string> equations = new List<string>();

        foreach (OfficeMath math in mathNodes)
        {
            // GetText provides a readable representation of the equation.
            string text = math.GetText().Trim();
            if (!string.IsNullOrEmpty(text))
                equations.Add(text);
        }

        // Write the extracted equations to a text file, one per line.
        string outputPath = "Equations.txt";
        File.WriteAllLines(outputPath, equations);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    // Helper that inserts an EQ field, converts it to a real OfficeMath node, and removes the field.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the field start's parent (the paragraph).
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // If conversion succeeded, replace the field with the OfficeMath node.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Insert a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
