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
            var folderPath = args[1];
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Папка с РПД '{folderPath}' не найдена");
                return;
            }
            if (!File.Exists(curriculumPath))
            {
                Console.WriteLine($"Файл с учебным планом '{curriculumPath}' не найден");
                return;
            }

            var files = Directory.EnumerateFiles(folderPath).ToList();
            var disciplines = new DocxCurriculum(curriculumPath).Disciplines;

            using var pattern = WordprocessingDocument.Open("./pattern.docx", false);
            using var QCReport = WordprocessingDocument.Create($"./Отчет РПД {new FileInfo(curriculumPath).Name}", WordprocessingDocumentType.Document);
            foreach (var part in pattern.Parts)
            {
                QCReport.AddPart(part.OpenXmlPart, part.RelationshipId);
            }

            var body = QCReport.MainDocumentPart.Document.Body;
            var table = body.Descendants<Table>().First();
            var analyzedProgramsCount = 0;
            var incorrectFormProgramsCount = 0;
            var incorrectValueFundProgramsCount = 0;
            var problems = new Dictionary<string, int>();

            foreach (var discipline in disciplines)
            {
                var programFileName = files.SingleOrDefault(f => f.Contains(discipline.Code));
                if (programFileName == null)
                {
                    continue;
                }

                try
                {
                    var (content, errors) = ProgramContentChecker.parseProgramFile(programFileName);

                    // sections check
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
                        if (!problems.ContainsKey(sectionNumber))
                        {
                            problems.Add(sectionNumber, 1);
                            continue;
                        }
                        ++problems[sectionNumber];
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
                        if (!problems.ContainsKey(sectionNumber))
                        {
                            problems.Add(sectionNumber, 1);
                            continue;
                        }
                        ++problems[sectionNumber];
                    }

                    var incorrectFormSections = missingSections.Append(emptySections);

                    // value fund check
                    var valueFundCheckResult = new StringBuilder();
                    valuationFund.Where(s => content.TryGetValue(s.Key, out var text) && text.Trim() != "").ToList()
                        .ForEach(s => valueFundCheckResult.Append(s.Value));

                    var row = new TableRow(
                        CreateTableCell($"[{discipline.Code}] {discipline.RussianName}"),
                        CreateTableCell($"{incorrectFormSections}"),
                        CreateTableCell($"{valueFundCheckResult}"),
                        CreateTableCell(""),
                        CreateTableCell(""));

                    table.Append(row);

                    // analytics
                    ++analyzedProgramsCount;
                    if (!string.IsNullOrEmpty(incorrectFormSections.ToString()))
                    {
                        ++incorrectFormProgramsCount;
                    }
                    if (!string.IsNullOrEmpty(valueFundCheckResult.ToString()))
                    {
                        ++incorrectValueFundProgramsCount;
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            PrintAnalytics(body, disciplines.Count, analyzedProgramsCount, incorrectFormProgramsCount, incorrectValueFundProgramsCount, problems);
        }

        private static void PrintAnalytics(Body body, int allProgramsCount, int analyzedProgramsCount, int incorrectFormProgramsCount, int incorrectValueFundProgramsCount, Dictionary<string, int> problems)
        {
            var table = body.Descendants<Table>().First();
            var paragraphs = table.ElementsAfter().Where(e => e is Paragraph).Skip(2).ToList();
            var info = new List<string>
            {
                $"--- {analyzedProgramsCount}/{allProgramsCount}",
                $"--- {incorrectFormProgramsCount}",
                $"--- {incorrectValueFundProgramsCount}"
            };

            for (var i = 0; i < 3; ++i)
            {
                paragraphs[i].AppendChild(new Run(new Text(info[i])));
            }

            foreach (var problem in problems.OrderByDescending(p => p.Value).ToList())
            {
                paragraphs[3].AppendChild(new Run(new Break()));
                paragraphs[3].AppendChild(new Run(new Text($"{problem.Key}  {problem.Value}")));
            }
        }

        private static readonly Dictionary<string, string> valuationFund = new()
        {
            {
                "3.1.3. Методика проведения текущего контроля " +
                    "успеваемости и промежуточной аттестации и критерии оценивания",
                "3.1.3, "
            },
            {
                "3.1.4. Методические материалы для проведения текущего контроля успеваемости и промежуточной" +
                " аттестации (контрольно-измерительные" +
                        " материалы, оценочные средства)",
                "3.1.4, "
            }
        };

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

        private static TableCell CreateTableCell(string text)
            => new(new Paragraph(new Run(new Text(text))));
    }
}
