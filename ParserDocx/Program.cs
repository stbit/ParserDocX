using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ParserDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            string projectDirectory = Directory.GetParent(Environment.CurrentDirectory).Parent.FullName;
            string distPath = Path.Combine(projectDirectory, "..", "documents", "dist");

            // Удаляем папку с результатами
            if (Directory.Exists(distPath))
            {
                try
                {
                    Directory.Delete(distPath, recursive: true);
                } catch
                {
                    Thread.Sleep(2000);
                    Directory.Delete(distPath, recursive: true);
                }
            }

            foreach (var filePath in Directory.GetFiles(Path.Combine(projectDirectory, "..", "documents"), "*.docx"))
            {
                Console.WriteLine(filePath);
                ConvertDocx.Parse(filePath).Save();
            }

            Console.WriteLine("finish!");

            Console.ReadLine();
        }
    }
}
