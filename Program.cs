using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using System.Threading;
using System.Windows;

namespace C_program
{
    public partial class SolidWorksMacro
    {
        public static void Main()
        {
            ModelDoc2 swModel;
            ModelDoc2 swDrawing;
            Frame pFrame;
            string fileName = null;
            string fileName1 = null;
            int errors = 0;
            int warnings = 0;
            int errors1 = 0;

            string fullPath = Path.GetFullPath("solidworks_conversion_program.exe");
            string filename = Path.GetFileName(fullPath);
            string root = fullPath.Replace(filename, "");

            string macrofilename = root + "step.swp";
            string macrofilename1 = root + "stl.swp";
            string macrofilename2 = root + "dxf.swp";
            string macrofilename3 = root + "pdf.swp";

            if (File.Exists(root + "temp_file_txt.txt") == false)
            {
                Console.WriteLine("brak pliku 'temp_file_txt'. Uruchom program do segregacji!!!");
                Console.WriteLine("Nacisnij dowolny klawisz aby zakończyć.");
                Console.ReadKey();
                System.Environment.Exit(0);
            }

            string[] lines = System.IO.File.ReadAllLines("temp_file_txt.txt");
            int line_number = lines.Length;
            string[] result = new string[line_number];
            int i = 0;
            Console.WriteLine("LISTA PLIKÓW DO KONWERSJI: \n");

            foreach (string line in lines)
            {
                Console.Write(i);
                Console.WriteLine(" - " + line + "\t");
                string[] separator = { " => " };
                string[] wynik = line.Split(separator, StringSplitOptions.None);
                result[i] = wynik[0];
                i = i + 1;
            }

            Console.WriteLine("\n\n ROZPOCZĘCIE KONWERSJI: \n");
            pFrame = (Frame)swapp.Frame();
            pFrame.KeepInvisible = true;

            foreach (string position in result)
            {
                Console.Write("\n" + position);
                fileName = position + ".sldprt";
                if (File.Exists(fileName) == true)
                {
                    swModel = (ModelDoc2)swapp.OpenDoc6(fileName, 1, 1, "", ref errors, ref warnings);
                    swapp.RunMacro2(macrofilename, "step", "main", 1, out errors1);
                    swapp.RunMacro2(macrofilename1, "stl", "main", 1, out errors1);
                    swapp.CreateNewWindow();
                }
                else
                {
                    fileName = position + ".sldasm";
                    Console.Write(" - to złożenie.");
                    swModel = (ModelDoc2)swapp.OpenDoc6(fileName, 2, 1, "", ref errors, ref warnings);
                    swapp.RunMacro2(macrofilename, "step", "main", 1, out errors1);
                    swapp.RunMacro2(macrofilename1, "stl", "main", 1, out errors1);
                    swapp.CreateNewWindow();
                }

                swapp.CloseAllDocuments(true);
                swapp.ExitApp();
                Thread.Sleep(1500);

                fileName1 = position + ".slddrw";
                if (File.Exists(fileName1) == true)
                {
                    swDrawing = (ModelDoc2)swapp.OpenDoc6(fileName1, 3, 1, "", ref errors, ref warnings);
                    swapp.RunMacro2(macrofilename2, "dxf", "main", 1, out errors1);
                    swapp.RunMacro2(macrofilename3, "pdf", "main", 1, out errors1);
                    pFrame.KeepInvisible = false;
                    swapp.Visible = true;
                    swapp.CreateNewWindow();
                    swapp.CloseAllDocuments(true);
                    swapp.ExitApp();
                    Thread.Sleep(1500);
                }
                else
                {
                    Console.Write("  ==> plik '.slddrw' z rysunkiem nie istnieje.");
                }

                Console.Write("  ==> wykonano!!");



            }


            swapp.ExitApp();

            using (StreamWriter sw = File.AppendText("temp_file_txt.txt"))
            {
                sw.WriteLine("ok");
            }

            Console.WriteLine("\n\n ZAKOŃCZONO KONWERSJE PLIKÓW.");
            Console.WriteLine("POCZEKAJ NA ZAKOŃCZENIE OPERACJI PRZENOSZENIA PLIKÓW DO ODPOWIEDNICH FOLDERÓW.");
        }
        public static SldWorks swapp = new SldWorks();
    }
}
