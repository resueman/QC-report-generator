using System;
using System.Collections.Generic;

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
                Console.WriteLine("dotnet run <учебный план 1> <папка с РПД учебного плана 1> <учебный план 2> <папка с РПД учебного плана 2> ... <учебный план N> <папка с РПД учебного плана N>");
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
