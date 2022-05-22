using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.IO.Compression;

namespace CreateListStringFromTextFile
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFilePath;
            string encodingsetting;
            if (args.Length == 2)
            {

                inputFilePath = args[0];
                encodingsetting = args[1];
            }
            else
            {
                Console.WriteLine("Must supply two args: path to text file, and encoding. For encoding use 1 for System default(ansi) codepage and 2 for UTF8");
                return;
            }


            StreamReader sr;

            // try to open file
            try
            {

                if (encodingsetting == "2")
                {
                    sr = new StreamReader(inputFilePath, Encoding.UTF8);
                }
                else
                {

                    sr = new StreamReader(inputFilePath, Encoding.Default);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Exception opening file: {0}", e.Message);
                return;
            }

            var list = new List<string>();
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                list.Add(line);

            }

            Console.WriteLine("Added {0} items.", list.Count);
            string fileName = inputFilePath.Substring(inputFilePath.LastIndexOf('\\') + 1);

            var binFormatter = new BinaryFormatter();
            var outfile = File.Open(fileName + ".bin", FileMode.Create);
            using (DeflateStream ds = new DeflateStream(outfile, CompressionLevel.Optimal))
            {
                binFormatter.Serialize(ds, list);
            }

        }
    }
}
