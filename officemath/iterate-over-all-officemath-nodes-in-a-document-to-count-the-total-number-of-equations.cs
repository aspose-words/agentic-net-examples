using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three simple equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");   // Fraction 1/2
        InsertEquation(builder, @"\r(2,x)");   // Square root of x
        InsertEquation(builder, @"\i");        // Integral symbol

        // Save the sample document.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        doc.Save(docPath);

        // Load the document (optional, demonstrates load workflow).
        Document loadedDoc = new Document(docPath);

        // Count all OfficeMath nodes that represent top‑level equations (MathObjectType.OMathPara).
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        int equationCount = 0;
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                equationCount++;
        }

        // Output the result.
        Console.WriteLine($"Total number of equations: {equationCount}");
    }

    // Helper method to insert an EQ field, convert it to OfficeMath, and clean up the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return to the field start's parent node.
        builder.MoveTo(field.Start.ParentNode);
        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }
        // Insert a paragraph break after each equation for readability.
        builder.InsertParagraph();
    }
}
