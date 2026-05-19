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

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertFieldEQ(builder, @"\f(1,2)");               // Simple fraction 1/2
        InsertFieldEQ(builder, @"\r(3,x)");               // Cube root of x
        InsertFieldEQ(builder, @"\i \su(n=1,5,n)");       // Integral with summation

        // Convert each inserted EQ field to a real OfficeMath node and remove the field.
        foreach (FieldEQ field in doc.Range.Fields.OfType<FieldEQ>())
        {
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field.
                field.Remove();
            }
        }

        // Extract all OfficeMath equations from the document.
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        string outputPath = "ExtractedEquations.txt";

        using (StreamWriter writer = new StreamWriter(outputPath, false))
        {
            foreach (OfficeMath om in officeMathNodes)
            {
                // GetText provides a readable representation of the equation.
                writer.WriteLine(om.GetText().Trim());
            }
        }

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the equations report file.");

        // (Optional) Save the document with the real OfficeMath nodes for inspection.
        doc.Save("SampleWithOfficeMath.docx");
    }

    // Helper method that inserts an EQ field, writes its arguments, and starts a new paragraph.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return the builder to the field start's parent (the paragraph) and start a new paragraph.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();
        return field;
    }
}
