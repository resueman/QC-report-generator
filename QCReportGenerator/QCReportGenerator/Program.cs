using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using CurriculumParser;
using System.IO;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace QCReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Добро пожаловать в QC Report Generator, утилиту для генерации отчета комиссии контроля качества");
                Console.WriteLine("Использование:");
                Console.WriteLine("dotnet run <номер рабочего плана и год> <папка с РПД>");
                return;
            }

            var curriculumPath = args[0];
            var rpdFolderPath = args[1];
            var programRpdsAnalyzer = new ProgramRpdsAnalyzer(curriculumPath, rpdFolderPath);
            var generator = new QcReportGenerator(programRpdsAnalyzer);
            generator.GenerateReport();
        }
    }
}
