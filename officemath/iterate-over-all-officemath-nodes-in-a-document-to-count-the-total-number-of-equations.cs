using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the sample document.
        const string filePath = "SampleEquations.docx";

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper that inserts an equation using the deterministic EQ‑field bootstrap workflow.
        void InsertEquation(string eqArgument)
        {
            // Insert an empty equation field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);

            // Move to the field separator node and write the EQ argument.
            builder.MoveTo(field.Separator);
            builder.Write(eqArgument);

            // Convert the field to a real OfficeMath node.
            FieldEQ eqField = (FieldEQ)field;
            OfficeMath officeMath = eqField.AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                Node fieldStart = eqField.Start;
                fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

                // Remove the original field so only the OfficeMath remains.
                eqField.Remove();
            }

            // Move to a new paragraph for the next equation.
            builder.Writeln();
        }

        // Insert several simple equations.
        InsertEquation(@"\f(1,2)");   // Fraction 1/2
        InsertEquation(@"\r(3,x)");   // Radical
        InsertEquation(@"\s(5)");     // Summation placeholder

        // Save the document.
        doc.Save(filePath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(filePath))
            throw new Exception("Failed to create the output document.");

        // Reload the document for counting.
        Document loadedDoc = new Document(filePath);

        // Get all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Count only top‑level equations (MathObjectType.OMathPara).
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the result.
        Console.WriteLine($"Total equations: {equationCount}");
    }
}
