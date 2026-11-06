using System;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        string[] progIds = { "Word.Application", "Excel.Application", "NonExistent.ProgID" };

        foreach (var progId in progIds)
        {
            bool isValid = IsProgIdRegistered(progId);
            Console.WriteLine($"{progId}: {(isValid ? "Valid" : "Invalid")}");

            if (isValid)
            {
                // Simulate safe insertion of the OLE object
                Console.WriteLine($"Simulating insertion of OLE object with ProgID '{progId}'.");
            }
        }
    }

    private static bool IsProgIdRegistered(string progId)
    {
        try
        {
            Type comType = Type.GetTypeFromProgID(progId);
            return comType != null;
        }
        catch (COMException)
        {
            return false;
        }
        catch (Exception)
        {
            return false;
        }
    }
}
