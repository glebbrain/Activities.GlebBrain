using System;
using System.Collections.Generic;
using System.Text.Json;
using Activities.GlebBrain;

namespace TestActivities
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Activities.GlebBrain.IsActiveWindow iaw = new IsActiveWindow();
            //iaw.Title = "Блокнот";
            iaw.ProcessName = "notepad.exe";
            iaw.Execute(1);

            if (iaw.IsActive)
            {
                Console.WriteLine("Все ок!");
            }
            else
            {
                Console.WriteLine(iaw.ExceptionMessage);
            }

            Activities.GlebBrain.SetActivesWindow saw = new SetActivesWindow();
            saw.Title = "Блокнот";
            saw.ProcessName = "notepad.exe";
            saw.Execute(1);
            
            if (saw.IsActive)
            {
                Console.WriteLine("Все ок!");
            }
            else
            {
                Console.WriteLine(saw.ExceptionMessage);
            }
        }
    }
}
