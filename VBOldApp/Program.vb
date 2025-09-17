Imports System

Module Program
    Sub Main(args As String())
        ' VB Old App - A simple demonstration of VB.NET features
        Console.WriteLine("=====================================")
        Console.WriteLine("        VB Old App v1.0")
        Console.WriteLine("  A Classic Visual Basic Application")
        Console.WriteLine("=====================================")
        Console.WriteLine()
        
        ' Display current date and time
        Console.WriteLine("Current Date and Time: " & DateTime.Now.ToString())
        Console.WriteLine()
        
        ' Simple calculator demonstration
        CalculatorDemo()
        
        ' String manipulation demonstration
        StringDemo()
        
        ' Wait for user input before closing
        Console.WriteLine()
        Console.Write("Press any key to exit...")
        Console.ReadKey()
    End Sub
    
    Sub CalculatorDemo()
        Console.WriteLine("=== Calculator Demo ===")
        
        Dim num1 As Double = 10
        Dim num2 As Double = 5
        
        Console.WriteLine("Number 1: " & num1)
        Console.WriteLine("Number 2: " & num2)
        Console.WriteLine()
        Console.WriteLine("Addition: " & num1 & " + " & num2 & " = " & (num1 + num2))
        Console.WriteLine("Subtraction: " & num1 & " - " & num2 & " = " & (num1 - num2))
        Console.WriteLine("Multiplication: " & num1 & " * " & num2 & " = " & (num1 * num2))
        Console.WriteLine("Division: " & num1 & " / " & num2 & " = " & (num1 / num2))
        Console.WriteLine()
    End Sub
    
    Sub StringDemo()
        Console.WriteLine("=== String Demo ===")
        
        Dim appName As String = "VB Old App"
        Dim version As String = "1.0"
        
        Console.WriteLine("Application Name: " & appName)
        Console.WriteLine("Version: " & version)
        Console.WriteLine("Length of app name: " & appName.Length)
        Console.WriteLine("Uppercase: " & appName.ToUpper())
        Console.WriteLine("Lowercase: " & appName.ToLower())
        Console.WriteLine("Full title: " & appName & " v" & version)
        Console.WriteLine()
    End Sub
End Module
