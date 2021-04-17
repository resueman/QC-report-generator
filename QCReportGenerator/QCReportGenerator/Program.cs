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
            if (args.Length < 2)
            {
                Console.WriteLine("Добро пожаловать в QC Report Generator, утилиту для генерации отчета комиссии контроля качества");
                Console.WriteLine("Использование:");
                Console.WriteLine("dotnet run <номер рабочего плана и год> <папка с РПД>");
                return;
            }

            var results = new List<ProgramRpdsAnalyzer>();
            for (var i = 0; i < args.Length; i += 2)
            {
                var curriculumPath = args[i];
                var rpdFolderPath = args[i + 1];
                var rpdsAnalysisResult = new ProgramRpdsAnalyzer(curriculumPath, rpdFolderPath);
                results.Add(rpdsAnalysisResult);
            }

            var generator = new QcReportGenerator(results);
            generator.GenerateReport();
        }
    }
}
