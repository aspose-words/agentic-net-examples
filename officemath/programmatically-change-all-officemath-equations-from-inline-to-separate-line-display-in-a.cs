using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathDisplayChanger
{
    public static void Main()
    {
        // Create a sample document with several inline equations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("This paragraph contains inline equations:");

        // Insert three inline equations using the deterministic EQ-field bootstrap workflow.
        InsertInlineEquation(builder, @"\f(1,2)");          // Simple fraction 1/2
        builder.Writeln(); // separate paragraph for readability
        InsertInlineEquation(builder, @"\r(3,x)");          // Cube root of x
        builder.Writeln();
        InsertInlineEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the initial document (optional, for inspection).
        string inlinePath = Path.Combine(Environment.CurrentDirectory, "ReportInline.docx");
        doc.Save(inlinePath, SaveFormat.Docx);

        // Change all top‑level OfficeMath equations from inline to display mode.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the modified document.
        string displayPath = Path.Combine(Environment.CurrentDirectory, "ReportDisplay.docx");
        doc.Save(displayPath, SaveFormat.Docx);

        // Simple validation that the output file was created.
        if (!File.Exists(displayPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, and sets it to inline initially.
    private static void InsertInlineEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();

            // Ensure the equation starts as inline (the state we will later change).
            officeMath.DisplayType = OfficeMathDisplayType.Inline;
        }
    }
}
