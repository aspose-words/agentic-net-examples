using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ExtractOfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few deterministic equations using the EQ field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");                 // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");                // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)");        // Integral with summation

        // -----------------------------------------------------------------
        // Extraction: collect all top‑level OfficeMath paragraphs (OMathPara)
        // -----------------------------------------------------------------
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        StringBuilder output = new StringBuilder();

        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // GetText() provides a readable representation of the equation.
                string text = om.GetText().Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    output.AppendLine(text);
                }
            }
        }

        // Write the extracted equations to a text file.
        const string outputFile = "Equations.txt";
        File.WriteAllText(outputFile, output.ToString());

        // Validate that the file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Failed to create the output file '{outputFile}'.");

        // Show the result in the console for verification.
        int equationCount = output.ToString()
                                   .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
                                   .Length;
        Console.WriteLine($"Extracted {equationCount} equations to '{outputFile}'.");
    }

    // Helper that creates a real OfficeMath node from an EQ field argument string.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that the EQ code is processed before conversion.
        field.Update();

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();

        // Move the builder after the inserted equation and start a new paragraph.
        builder.MoveTo(officeMath);
        builder.InsertParagraph();

        return officeMath;
    }
}
