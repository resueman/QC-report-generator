using CurriculumParser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace QCReportGenerator
{
    class ProgramRpdsAnalyzer
    {
        public string CurriculumPath { get; private set; }

        public string RpdFolderPath { get; private set; }

        public int ExpectedProgramsCount { get; private set; }

        public int ActualProgramsCount { get; private set; }
        
        public int IncorrectFormProgramsCount { get; private set; }
        
        public int IncorrectValueFundProgramsCount { get; private set; }
        
        public Dictionary<string, int> RpdProblemsFrequency { get; private set; }
        
        public Dictionary<IgnoreReasonType, List<string>> IgnoredRpd { get; private set; }

        public List<(Discipline Discipline, string FormMismatchSections, string ValueFundCheckResult)> Results { get; private set; }

        private static readonly Dictionary<string, string> valuationFund = new()
        {
            {
                "3.1.3. Методика проведения текущего контроля " +
            "успеваемости и промежуточной аттестации и критерии оценивания",
                "3.2.4, "
            },
            {
                "3.1.4. Методические материалы для проведения текущего контроля успеваемости и промежуточной" +
            " аттестации (контрольно-измерительные материалы, оценочные средства)",
                "3.2.1, 3.2.2, 3.2.3, "
            }
        };

        public ProgramRpdsAnalyzer(string curriculumPath, string rpdFolderPath)
        {
            if (!Directory.Exists(rpdFolderPath))
            {
                throw new Exception($"Папка с РПД '{rpdFolderPath}' не найдена");
            }
            if (!File.Exists(curriculumPath))
            {
                throw new Exception($"Файл с учебным планом '{curriculumPath}' не найден");
            }

            CurriculumPath = curriculumPath;
            RpdFolderPath = rpdFolderPath;

            Results = new List<(Discipline Discipline, string FormMismatchSections, string ValueFundCheckResult)>();
            RpdProblemsFrequency = new Dictionary<string, int>();
            IgnoredRpd = new Dictionary<IgnoreReasonType, List<string>>
            {
                { IgnoreReasonType.NotFound, new List<string>() },
                { IgnoreReasonType.ParsingProblems, new List<string>() }
            };

            Analyze();
        }

        private void Analyze()
        {
            var files = Directory.EnumerateFiles(RpdFolderPath).ToList();
            var disciplines = new DocxCurriculum(CurriculumPath).Disciplines;
            ExpectedProgramsCount = disciplines.Count;
            foreach (var discipline in disciplines)
            {
                var programFileName = files.SingleOrDefault(f => f.Contains(discipline.Code));
                if (programFileName == null)
                {
                    IgnoredRpd[IgnoreReasonType.NotFound].Add($"{discipline.Code} {discipline.RussianName}");
                    continue;
                }

                List<string> errors;
                Dictionary<string, string> content;
                try
                {
                    var (c, e) = ProgramContentChecker.parseProgramFile(programFileName);
                    ++ActualProgramsCount;
                    errors = e.ToList();
                    content = c.ToDictionary(c => c.Key, c => c.Value);
                }
                catch (Exception)
                {
                    IgnoredRpd[IgnoreReasonType.ParsingProblems].Add(programFileName);
                    continue;
                }

                var formMismatchSections = GetEstablishedFormMismatchErrors(content, errors);
                var valueFundCheckResult = GetValueFundCheckErrors(content);
                Results.Add((discipline, formMismatchSections, valueFundCheckResult));
            }
        }

        private string GetValueFundCheckErrors(Dictionary<string, string> content)
        {
            var valueFundCheckResult = new StringBuilder();
            valuationFund.Where(s => content.TryGetValue(s.Key, out var text) && text.Trim() != "").ToList()
                .ForEach(s => valueFundCheckResult.Append(s.Value));

            if (!string.IsNullOrEmpty(valueFundCheckResult.ToString()))
            {
                ++IncorrectValueFundProgramsCount;
            }

            return valueFundCheckResult.ToString();
        }

        private string GetEstablishedFormMismatchErrors(Dictionary<string, string> content, List<string> errors)
        {
            // get and count missing sections
            var missingSections = new StringBuilder();
            foreach (var error in errors)
            {
                var sectionNumber = GetMissingSectionNumber(error);
                if (string.IsNullOrEmpty(sectionNumber))
                {
                    continue;
                }
                missingSections.Append($"{sectionNumber}, ");
                if (!RpdProblemsFrequency.ContainsKey(sectionNumber))
                {
                    RpdProblemsFrequency.Add(sectionNumber, 1);
                    continue;
                }
                ++RpdProblemsFrequency[sectionNumber];
            }

            // get and count empty sections
            var emptySections = new StringBuilder();
            foreach (var section in content.Where(s => s.Value == "").ToList())
            {
                var sectionNumber = GetEmptySectionNumber(section.Key);
                if (string.IsNullOrEmpty(sectionNumber))
                {
                    continue;
                }
                emptySections.Append($"{sectionNumber}, ");
                if (!RpdProblemsFrequency.ContainsKey(sectionNumber))
                {
                    RpdProblemsFrequency.Add(sectionNumber, 1);
                    continue;
                }
                ++RpdProblemsFrequency[sectionNumber];
            }

            var incorrectFormSections = missingSections.Append(emptySections);

            if (!string.IsNullOrEmpty(incorrectFormSections.ToString()))
            {
                ++IncorrectFormProgramsCount;
            }

            return incorrectFormSections.ToString();
        }

        private static string GetMissingSectionNumber(string errorMessage)
        {
            var match = Regex.Match(errorMessage, @"'([0-9\.]+)',?.*");
            return match.Success ? FormatSectionNumber(match.Groups[1].Value) : "";
        }

        private static string GetEmptySectionNumber(string s)
        {
            var match = Regex.Match(s, @"^([0-9\.]+).*");
            return match.Success ? FormatSectionNumber(match.Groups[1].Value) : "";
        }

        private static string FormatSectionNumber(string number)
            => number.EndsWith('.')
                ? number.Substring(0, number.Length - 1)
                : number;
    }
}
