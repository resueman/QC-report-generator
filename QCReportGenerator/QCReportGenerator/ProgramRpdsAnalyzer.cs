using CurriculumParser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace QCReportGenerator
{
    /// <summary>
    /// Анализирует рабочие программы дисциплин, которые будут проводится в будущем учебном году у конкретного направления обучения
    /// </summary>
    public class ProgramRpdsAnalyzer
    {
        public DocxCurriculum Curriculum { get; private set; }

        /// <summary>
        /// Курс бакалавриата, на котором будут учиться студенты в будущем учебном году
        /// </summary>
        public int Course { get; private set; }

        public string CurriculumPath { get; private set; }

        public string RpdFolderPath { get; private set; }

        /// <summary>
        /// Количество дисциплин для изучения на будущий учебный год, указанное в учебном плане
        /// </summary>
        public int ExpectedProgramsCount { get; private set; }
        
        /// <summary>
        /// Количество дисциплин в таблице, тех, при обработке которых не возникло исключений
        /// </summary>
        public int ActualProgramsCount { get; private set; }
        
        /// <summary>
        /// Количество РПД не соответсвующих установленной форме (столбец 2 таблицы)
        /// </summary>
        public int IncorrectFormProgramsCount { get; private set; }
        
        /// <summary>
        /// Количество РПД, в которых ФОС не соответствует Положению (столбец 3 таблицы)  
        /// </summary>
        public int IncorrectValueFundProgramsCount { get; private set; }
        
        /// <summary>
        /// Общее число ошибок для для секций (только секции, в которых есть хотя бы одна ошибка)
        /// </summary>
        public Dictionary<string, int> RpdSectionProblemsFrequency { get; private set; }
        
        /// <summary>
        /// РПД, которые не были обработаны с информацией о причине
        /// </summary>
        public Dictionary<IgnoreReasonType, List<string>> IgnoredRpd { get; private set; }

        /// <summary>
        /// Результат анализа РПД программы обучения
        /// </summary>
        public List<(Discipline Discipline, string FormMismatchSections, string ValueFundCheckResult)> Results { get; private set; }

        /// <summary>
        /// Фонд оценочных средств. Ключ -- секция РПД, описывающая ФОС, значение -- пункт Положения, выполнимость 
        /// которого определяется содержанием секции РПД 
        /// </summary>
        private static readonly Dictionary<string, string> valuationFund = new()
        {
            {
                "3.1.4. Методические материалы для проведения текущего контроля успеваемости и промежуточной" +
                    " аттестации (контрольно-измерительные материалы, оценочные средства)",
                "3.2.3, "
            },
            {
                "3.1.3. Методика проведения текущего контроля " +
                    "успеваемости и промежуточной аттестации и критерии оценивания",
                "3.2.4, "
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
            RpdSectionProblemsFrequency = new Dictionary<string, int>();
            IgnoredRpd = new Dictionary<IgnoreReasonType, List<string>>
            {
                { IgnoreReasonType.NotFound, new List<string>() },
                { IgnoreReasonType.ParsingProblems, new List<string>() },
                { IgnoreReasonType.TwoRpdsInFolder, new List<string>() }
            };

            Curriculum = new DocxCurriculum(CurriculumPath);
            var curriculumYear = Curriculum.CurriculumCode.Substring(0, 2);
            Course = 1 + DateTime.Now.Year - 2000 - int.Parse(curriculumYear);
            Analyze();
        }

        /// <summary>
        /// Анализирует РПД
        /// </summary>
        private void Analyze()
        {
            var files = Directory.EnumerateFiles(RpdFolderPath).ToList();
            var disciplines = Curriculum.Disciplines
                .Where(d => d.Implementations.Select(i => i.Semester).Contains(Course * 2 - 1)
                    || d.Implementations.Select(i => i.Semester).Contains(Course * 2))
                .ToList();

            ExpectedProgramsCount = disciplines.Count;
            foreach (var discipline in disciplines)
            {
                var programFileName = "";
                try
                {
                    programFileName = files.SingleOrDefault(f => f.Contains(discipline.Code));
                    if (programFileName == null)
                    {
                        IgnoredRpd[IgnoreReasonType.NotFound].Add($"{discipline.Code} {discipline.RussianName}");
                        continue;
                    }
                }
                catch (InvalidOperationException)
                {
                    var names = "";
                    files.Where(f => f.Contains(discipline.Code)).ToList().ForEach(f => names += $"{f} ");
                    IgnoredRpd[IgnoreReasonType.TwoRpdsInFolder].Add($"{names}");
                    continue;
                }

                try
                {
                    var (c, e) = ProgramContentChecker.parseProgramFile(programFileName);
                    ++ActualProgramsCount;
                    var content = c.ToDictionary(c => c.Key, c => c.Value);
                    var formMismatchSections = GetEstablishedFormMismatchErrors(content, e.ToList());
                    var valueFundCheckResult = GetValueFundCheckErrors(c);
                    Results.Add((discipline, formMismatchSections, valueFundCheckResult));
                }
                catch (Exception)
                {
                    IgnoredRpd[IgnoreReasonType.ParsingProblems].Add(programFileName);
                    continue;
                }
            }
        }

        /// <summary>
        /// Проверяет соответсвие РПД установленной форме (столбец 1 таблицы)
        /// </summary>
        /// <param name="content">Содержание РПД</param>
        /// <param name="errors">Ошибки возникшие при разборе РПД</param>
        /// <returns>Пустую строку, если РПД соответствует установленной форме, иначе пункты, которые не соответсвуют</returns>
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
                if (!RpdSectionProblemsFrequency.ContainsKey(sectionNumber))
                {
                    RpdSectionProblemsFrequency.Add(sectionNumber, 1);
                    continue;
                }
                ++RpdSectionProblemsFrequency[sectionNumber];
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
                if (!RpdSectionProblemsFrequency.ContainsKey(sectionNumber))
                {
                    RpdSectionProblemsFrequency.Add(sectionNumber, 1);
                    continue;
                }
                ++RpdSectionProblemsFrequency[sectionNumber];
            }

            var incorrectFormSections = missingSections.Append(emptySections);

            if (!string.IsNullOrEmpty(incorrectFormSections.ToString()))
            {
                ++IncorrectFormProgramsCount;
            }

            return FormatSequenceOfSections(incorrectFormSections.ToString());
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
                ? number[0..^1]
                : number;

        /// <summary>
        /// Оценивает соответсвие ФОС Положению (столбец 2 таблицы)
        /// </summary>
        /// <param name="c">Содержание РПД</param>
        /// <returns>Пустую строку, если ФОС рабочей программы соответсвует положению, иначе пункты Положения, которые не выполнены</returns>
        private string GetValueFundCheckErrors(Microsoft.FSharp.Collections.FSharpMap<string, string> c)
        {
            var valueFundCheckResult = new StringBuilder();
            var competencesError = ProgramContentChecker.shallContainCompetences(c).ToList();
            if (competencesError.Count != 0)
            {
                valueFundCheckResult.Append("3.2.1, 3.2.2, ");
            }

            var content = c.ToDictionary(c => c.Key, c => c.Value);
            valuationFund
                .Where(s => !content.TryGetValue(s.Key, out var text) || text.Trim() == "")
                .ToList()
                .ForEach(s => valueFundCheckResult.Append(s.Value));

            if (!string.IsNullOrEmpty(valueFundCheckResult.ToString()))
            {
                ++IncorrectValueFundProgramsCount;
            }

            return FormatSequenceOfSections(valueFundCheckResult.ToString());
        }

        private static string FormatSequenceOfSections(string sequence) 
            => sequence.EndsWith(", ") ? sequence[0..^2] : sequence;
    }
}
