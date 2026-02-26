// ALL ATTEMPTS FAILED. Below is the last generated code.

using Aspose.Words;
using Aspose.Words.Lists;
using System.Drawing.Printing;
using System.Windows.Forms;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Create a bulleted list ----------
        // Use AddSingleLevelList to create a single‑level list based on a template.
        List bulletedList = doc.Lists.AddSingleLevelList(ListTemplate.BulletCircle);
        builder.Writeln("Bulleted list starts below:");
        builder.ListFormat.List = bulletedList;   // Apply the list to subsequent paragraphs.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();       // End the list.

        // ---------- Create a numbered list ----------
        // Use Add to create a multi‑level list based on a predefined template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberUppercaseLetterDot);
        builder.Writeln("Numbered list starts below:");
        builder.ListFormat.List = numberedList;   // Apply the list.
        builder.Writeln("Item A");
        builder.Writeln("Item B");
        builder.ListFormat.RemoveNumbers();       // End the list.

        // ---------- Print the document programmatically ----------
        // Sends the document to the default printer without user interaction.
        doc.Print();

        // ---------- Print the document via a print dialog ----------
        // Allows the user to select printer settings before printing.
        using (PrintDialog printDialog = new PrintDialog())
        {
            printDialog.AllowSomePages = true;
            printDialog.PrinterSettings = new PrinterSettings();

            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                // Print using the printer settings chosen by the user.
                doc.Print(printDialog.PrinterSettings);
            }
        }
    }
}
