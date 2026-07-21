using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathEnumerationExample
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file that contains a few OfficeMath equations.
        // -----------------------------------------------------------------
        string samplePath = "SampleWithEquations.docx";
        CreateSampleDocument(samplePath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document that was just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(samplePath);

        // -----------------------------------------------------------------
        // 3. Enumerate all OfficeMath nodes in the document.
        // -----------------------------------------------------------------
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        Console.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");

        for (int i = 0; i < mathNodes.Count; i++)
        {
            OfficeMath om = (OfficeMath)mathNodes[i];
            Console.WriteLine($"--- Equation #{i + 1} ---");
            Console.WriteLine($"MathObjectType : {om.MathObjectType}");
            Console.WriteLine($"DisplayType    : {om.DisplayType}");
            Console.WriteLine($"Justification  : {om.Justification}");
            Console.WriteLine($"Text (for reference) : {om.GetText().Trim()}");
        }

        // -----------------------------------------------------------------
        // 4. Optionally, write a simple report file with the enumeration results.
        // -----------------------------------------------------------------
        string reportPath = "OfficeMathReport.txt";
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine($"Total OfficeMath nodes found: {mathNodes.Count}");
            for (int i = 0; i < mathNodes.Count; i++)
            {
                OfficeMath om = (OfficeMath)mathNodes[i];
                writer.WriteLine($"--- Equation #{i + 1} ---");
                writer.WriteLine($"MathObjectType : {om.MathObjectType}");
                writer.WriteLine($"DisplayType    : {om.DisplayType}");
                writer.WriteLine($"Justification  : {om.Justification}");
                writer.WriteLine($"Text : {om.GetText().Trim()}");
            }
        }

        // Verify that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // The program finishes here without waiting for user input.
    }

    // -----------------------------------------------------------------
    // Helper method that creates a DOCX with a few deterministic equations
    // using the EQ-field bootstrap workflow.
    // -----------------------------------------------------------------
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation: a simple fraction 1/2
        InsertEquation(builder, @"\f(1,2)");

        // Second equation: cube root of x
        InsertEquation(builder, @"\r(3,x)");

        // Third equation: an integral with limits
        InsertEquation(builder, @"\i \su(n=1,5,n)");

        // Save the document so it can be reloaded later.
        doc.Save(filePath);
    }

    // -----------------------------------------------------------------
    // Inserts an EQ field, converts it to a real OfficeMath node,
    // and removes the original field.
    // -----------------------------------------------------------------
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Ensure conversion succeeded before inserting.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field from the document.
            field.Remove();
        }

        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
