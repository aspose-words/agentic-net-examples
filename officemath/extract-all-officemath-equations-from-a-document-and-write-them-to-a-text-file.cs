using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Paths for the sample document and the output text file.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string txtPath = Path.Combine(Directory.GetCurrentDirectory(), "Equations.txt");

        // 1. Create a sample DOCX with a few OfficeMath equations using the deterministic EQ-field bootstrap workflow.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // First equation.
        builder.Writeln("First paragraph before equation.");
        InsertEquation(builder, @"\f(1,2)");
        builder.Writeln();

        // Second equation.
        builder.Writeln("Second paragraph before equation.");
        InsertEquation(builder, @"\f(3,4)");
        builder.Writeln();

        // Third equation.
        builder.Writeln("Third paragraph before equation.");
        InsertEquation(builder, @"\f(5,6)");
        builder.Writeln();

        // Save the sample document.
        sampleDoc.Save(docPath, SaveFormat.Docx);

        // Verify the sample document was created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create the sample DOCX file.", docPath);

        // 2. Load the document and extract all top‑level OfficeMath equations.
        Document loadDoc = new Document(docPath);
        NodeCollection mathNodes = loadDoc.GetChildNodes(NodeType.OfficeMath, true);

        using (StreamWriter writer = new StreamWriter(txtPath, false))
        {
            int equationIndex = 1;
            foreach (OfficeMath math in mathNodes)
            {
                if (math.MathObjectType == MathObjectType.OMathPara)
                {
                    string equationText = math.GetText().Trim();
                    writer.WriteLine($"Equation {equationIndex}: {equationText}");
                    equationIndex++;
                }
            }
        }

        // Validate that the text file was created and contains at least one equation.
        if (!File.Exists(txtPath))
            throw new FileNotFoundException("Failed to create the equations text file.", txtPath);

        string[] lines = File.ReadAllLines(txtPath);
        if (lines.Length == 0)
            throw new InvalidOperationException("No equations were written to the text file.");

        // Output the extracted equations to the console for verification (non‑interactive).
        foreach (string line in lines)
        {
            Console.WriteLine(line);
        }
    }

    // Helper method that creates a real OfficeMath node from an EQ field and inserts it into the document.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an empty Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ eqField = field as FieldEQ;
        if (eqField == null)
            throw new InvalidOperationException("Failed to create a FieldEQ.");

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(eqField.Separator);
        builder.Write(eqArgument);

        // Update the field so that Aspose.Words converts it to OfficeMath.
        eqField.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node fieldStart = eqField.Start;
        CompositeNode parent = fieldStart.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Field start does not have a composite parent.");

        parent.InsertBefore(officeMath, fieldStart);

        // Remove the original field.
        eqField.Remove();
    }
}
