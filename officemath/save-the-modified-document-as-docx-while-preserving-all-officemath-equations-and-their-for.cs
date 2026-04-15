using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Document with OfficeMath equations");
        builder.ParagraphFormat.ClearFormatting();

        // Insert several equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");          // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");          // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Modify formatting of top‑level OfficeMath nodes (display on its own line and left‑justified).
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the document as DOCX, preserving all OfficeMath equations and their formatting.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedDocument.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Optionally, output the count of top‑level equations for verification.
        int equationCount = doc.GetChildNodes(NodeType.OfficeMath, true)
                               .Count(node => ((OfficeMath)node).MathObjectType == MathObjectType.OMathPara);
        Console.WriteLine($"Document saved to '{outputPath}'. Top‑level equations count: {equationCount}");
    }

    // Helper method that inserts an EQ field, converts it to a real OfficeMath node,
    // removes the original field, and returns the created OfficeMath object.
    private static OfficeMath InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return to the paragraph containing the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field from the document.
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();

        return officeMath;
    }
}
