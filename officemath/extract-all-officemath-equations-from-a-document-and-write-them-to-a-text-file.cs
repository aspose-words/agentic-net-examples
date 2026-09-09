using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ExtractOfficeMath
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertFieldEQ(builder, @"\f(1,2)");          // Fraction 1/2
        InsertFieldEQ(builder, @"\r(3,x)");          // Cube root of x
        InsertFieldEQ(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Convert all inserted EQ fields to real OfficeMath objects.
        foreach (FieldEQ fieldEq in doc.Range.Fields.OfType<FieldEQ>().ToList())
        {
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and then remove the field.
                fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
                fieldEq.Remove();
            }
        }

        // Save the sample document (optional, just to demonstrate saving).
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Extract all OfficeMath equations from the document.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        string[] equations = mathNodes
            .Cast<OfficeMath>()
            .Select(m => m.GetText().Trim())
            .Where(t => !string.IsNullOrEmpty(t))
            .ToArray();

        // Write the extracted equations to a text file, one per line.
        string txtPath = "Equations.txt";
        File.WriteAllLines(txtPath, equations);

        // Validate that the output files were created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException($"Document file was not created: {docPath}");
        if (!File.Exists(txtPath))
            throw new FileNotFoundException($"Text report file was not created: {txtPath}");
    }

    // Helper that inserts an EQ field, writes its arguments, and moves the builder to a new paragraph.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write the EQ arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);
        // Return to the field start position.
        builder.MoveTo(field.Start.ParentNode);
        // Start a new paragraph after the field.
        builder.InsertParagraph();
        return field;
    }
}
