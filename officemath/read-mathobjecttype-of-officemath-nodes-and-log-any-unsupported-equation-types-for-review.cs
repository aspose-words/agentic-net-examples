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
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several equations using the deterministic EQ‑field bootstrap workflow.
        // Each equation uses a different switch to generate varied MathObjectTypes.
        InsertEquation(builder, @"\f(1,2)");                 // Fraction
        InsertEquation(builder, @"\r(3,x)");                 // Radical
        InsertEquation(builder, @"\a \co2 (a,b,c,d)");      // Array (matrix‑like)
        InsertEquation(builder, @"\i \su(n=1,5,n)");        // Integral with summation
        InsertEquation(builder, @"\s \up8(Sup) \s \do8(Sub)"); // Subscript/Superscript

        // Save the sample document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "OfficeMathSample.docx");
        doc.Save(docPath);

        // Enumerate all OfficeMath nodes in the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        using (StreamWriter reportWriter = new StreamWriter(Path.Combine(artifactsDir, "UnsupportedMathTypes.txt")))
        {
            int index = 0;
            foreach (OfficeMath officeMath in mathNodes)
            {
                // Log any MathObjectType that is not a top‑level paragraph (OMathPara).
                if (officeMath.MathObjectType != MathObjectType.OMathPara)
                {
                    string message = $"Unsupported MathObjectType: {officeMath.MathObjectType} (Node index {index})";
                    Console.WriteLine(message);
                    reportWriter.WriteLine(message);
                }
                index++;
            }
        }

        // Verify that the report file was created.
        string reportPath = Path.Combine(artifactsDir, "UnsupportedMathTypes.txt");
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // Optionally, save the document again to demonstrate that the workflow completes.
        string finalDocPath = Path.Combine(artifactsDir, "OfficeMathSample_Final.docx");
        doc.Save(finalDocPath);
    }

    // Helper that inserts an EQ field, converts it to OfficeMath, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the switch/arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqSwitch);

        // Move back to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
