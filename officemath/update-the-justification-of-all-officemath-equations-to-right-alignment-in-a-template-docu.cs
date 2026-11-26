using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Sample document with OfficeMath equations:");

        // Helper method that creates a top‑level OfficeMath equation using the EQ‑field bootstrap pattern.
        void InsertEquation(string eqSwitch)
        {
            // Insert an empty Equation field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ eqField = field as FieldEQ;
            if (eqField == null)
                throw new InvalidOperationException("Failed to cast the inserted field to FieldEQ.");

            // Write the EQ argument (the equation switch) into the field separator.
            builder.MoveTo(eqField.Separator);
            builder.Write(eqSwitch);

            // Update the field so that Aspose.Words can convert it to OfficeMath.
            eqField.Update();

            // Convert the field to an OfficeMath object.
            OfficeMath officeMath = eqField.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

            // Insert the OfficeMath node before the field start.
            Node parent = eqField.Start.ParentNode;
            if (parent is CompositeNode compositeParent)
            {
                compositeParent.InsertBefore(officeMath, eqField.Start);
            }
            else
            {
                throw new InvalidOperationException("Parent node is not a CompositeNode.");
            }

            // Remove the original field, leaving only the real OfficeMath node.
            eqField.Remove();
        }

        // Insert a few simple equations, each on its own paragraph.
        InsertEquation(@"\f(1,2)"); // Fraction 1/2
        builder.Writeln(); // Blank line between equations
        InsertEquation(@"\r(3,x)"); // Cube root of x
        builder.Writeln();
        InsertEquation(@"\s(2,a+b)"); // Summation from 2 to (a+b)

        // Save the template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        doc.Save(templatePath, SaveFormat.Docx);

        // Load the document again (simulating a separate processing step).
        Document loadedDoc = new Document(templatePath);

        // Update justification of all top‑level OfficeMath equations to right alignment.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.Justification = OfficeMathJustification.Right;
            }
        }

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Validation: ensure all top‑level OfficeMath nodes have right justification.
        Document verifyDoc = new Document(outputPath);
        NodeCollection verifyMathNodes = verifyDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in verifyMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara && om.Justification != OfficeMathJustification.Right)
                throw new InvalidOperationException("An OfficeMath node does not have right justification as expected.");
        }

        // Program completed successfully.
    }
}
