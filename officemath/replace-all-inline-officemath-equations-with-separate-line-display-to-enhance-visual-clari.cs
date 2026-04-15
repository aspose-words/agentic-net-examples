using System;
using System.IO;
using System.Linq;
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

        // Insert a few inline equations using EQ fields.
        // Equation 1: a simple fraction \f(1,2)
        InsertInlineEquation(builder, @"\f(1,2)");
        // Equation 2: a radical \r(3,x)
        InsertInlineEquation(builder, @"\r(3,x)");
        // Equation 3: an integral with summation \i \su(n=1,5,n)
        InsertInlineEquation(builder, @"\i \su(n=1,5,n)");

        // Convert all EQ fields to real OfficeMath objects.
        var eqFields = doc.Range.Fields.OfType<FieldEQ>().ToList();
        foreach (FieldEQ field in eqFields)
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

        // Change all top‑level inline OfficeMath equations to display (separate line) mode.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath math in mathNodes.OfType<OfficeMath>())
        {
            if (math.MathObjectType == MathObjectType.OMathPara && math.DisplayType == OfficeMathDisplayType.Inline)
            {
                math.DisplayType = OfficeMathDisplayType.Display;
                math.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method to insert an inline EQ field and convert it to OfficeMath later.
    private static void InsertInlineEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return to the field start's parent (the paragraph) and continue.
        builder.MoveTo(field.Start.ParentNode);
        // Insert a space after the equation to keep it inline with surrounding text.
        builder.Write(" ");
    }
}
