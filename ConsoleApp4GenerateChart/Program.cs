using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using MiscUtils;
using CommandLine.Text;
using CommandLine;
using System.Reflection;
using System.Windows.Forms;

namespace ConsoleApp4GenerateChart
{
    class Program
    {
        public class Options
        {
            [Option('i', "input", Required = false, HelpText = "Input path to parse file.")]
            public string InputPath { get; set; }

            [Option('o', "output", HelpText = "The full path name of output pdf files.")]
            public string OutputPath { get; set; }

            [Option('d', "debug", HelpText = "Log to files.")]
            public bool DebugFlag { get; set; }

            [Option('D', "date", HelpText = "Date want to parse.")]
            public string Date { get; set; }


            private string GetAssemblyCopyright()
            {
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                    return "";
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }


            [HelpOption]
            public string GetUsage()
            {
                var help = new HelpText
                {
                    Heading = new HeadingInfo(Application.ProductName, Application.ProductVersion),
                    Copyright = GetAssemblyCopyright(),
                    AdditionalNewLineAfterOption = true,
                    AddDashesToOption = true
                };


                help.AddPreOptionsLine(" ");
                help.AddPreOptionsLine(string.Format("Usage: {0} [options]", Path.GetFileName(Application.ExecutablePath)));

                help.AddPreOptionsLine(" ");
                help.AddPreOptionsLine("options:");
                help.AddOptions(this);
                return help;
            }
        }

        static void Main(string[] args)
        {


            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                
                if (options.DebugFlag)
                    Log.FileOutput(true, Path.Combine(Directory.GetCurrentDirectory(), "Log"), "genReportLog");

                string date = options.Date;
                if (string.IsNullOrEmpty(options.Date))
                {
                    //reqular parser
                    date = DateTime.Now.AddMonths(-1).ToString("MM/dd/yyyy");
                }

                MyReportOutputter mro = new MyReportOutputter();

                if (options.InputPath != null && options.OutputPath != null)
                    mro.StartBy(date, options.InputPath, options.OutputPath);

                else
                    mro.StartBy(options.Date);
            }
            else
            {
                options.GetUsage();
                return;
            }
            return;
        }

    }

}
