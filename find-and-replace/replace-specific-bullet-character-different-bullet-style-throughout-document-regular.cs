using System;
using System.Text.RegularExpressions;
using Aspose.Words;

class BulletReplacer
{
    /// <summary>
    /// Replaces all occurrences of a specific bullet character (•) with another bullet style (◦) in the given document.
    /// </summary>
    /// <param name="inputPath">Full path to the source .docx file.</param>
    /// <param name="outputPath">Full path where the modified document will be saved.</param>
    public static void ReplaceBullet(string inputPath, string outputPath)
    {
        // Load the existing document.
        Document doc = new Document(inputPath);

        // Regular expression that matches the bullet character (Unicode U+2022).
        Regex bulletPattern = new Regex("\u2022");

        // New bullet character to replace with (Unicode U+25E6 – white bullet).
        string replacementBullet = "\u25E6";

        // Replace every match of the bullet pattern throughout the document's range.
        doc.Range.Replace(bulletPattern, replacementBullet);

        // Save the updated document.
        doc.Save(outputPath);
    }
}

class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: BulletReplacer <input.docx> <output.docx>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            BulletReplacer.ReplaceBullet(inputPath, outputPath);
            Console.WriteLine($"Bullet characters replaced successfully. Saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
