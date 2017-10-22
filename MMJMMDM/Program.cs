using System;
using System.IO;
using System.Linq;

namespace MMJMMDM
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("*** Welcome to MMJMMDM...");
            Console.WriteLine("*** Please enter the directory you wish to Magically Data Munch (or leave blank to CANCEL)...");
            var directory = Console.ReadLine();

            if (directory == null || directory.Trim() == string.Empty)
            {
                Environment.Exit(0);
            }

            var hacker = new HackIt();
            var files = Directory.GetFiles(directory).ToList().Where(x => x.EndsWith(".xlsx")).Select(x => new FileInfo(x)).ToList();

            Console.WriteLine("*** MUNCHING {0} XLSX files...", files.Count());

            foreach (var file in files)
            {
                Console.WriteLine(file.FullName);
            }

            var results = hacker.GetStringsFromFiles(files);

            var newFileName = string.Format("MDM_{0}.csv", (Int32) (DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds);
            var newFilePath = string.Format("{0}\\{1}", directory, newFileName);
            var fi = new FileInfo(newFilePath);

            using (StreamWriter sw = fi.CreateText())
            {
                foreach (var result in results)
                {
                    sw.WriteLine(result);
                }
            }

            Console.WriteLine();

            Console.WriteLine("*** FINISHED!! The text is now in file {0} and can be imported into SPSS", newFilePath);
            Console.WriteLine("*** Please press any key to close MMJMMDM...");

            Console.ReadKey();
            Environment.Exit(0);
        }
    }
}
