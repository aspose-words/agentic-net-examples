using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "OfficeMathCount.docx");

        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple EQ field arguments that will be converted to real OfficeMath objects.
        string[] eqArguments = { @"\f(1,2)", @"\r(3,x)", @"\i" };

        foreach (string args in eqArguments)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

            // Write the EQ arguments into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(args);

            // Return the builder to the field start's parent (the paragraph).
            builder.MoveTo(field.Start.ParentNode);

            // Convert the field to an OfficeMath object.
            OfficeMath officeMath = field.AsOfficeMath();

            // Replace the field with the real OfficeMath node.
            if (officeMath != null)
            {
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                field.Remove();
            }

            // Start a new paragraph for the next equation.
            builder.Writeln();
        }

        // Save the document.
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Count top‑level OfficeMath nodes (equations) in the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the result.
        Console.WriteLine($"Total equations (top‑level OfficeMath): {equationCount}");
    }
}
