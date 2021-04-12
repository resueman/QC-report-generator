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
                    // get missing sections
                    var missingSections = new StringBuilder();
                    errors.ToList()
                        .ForEach(e => missingSections.Append(GetMissingSectionNumber(e)));

                    // get empty sections
                    var emptySections = new StringBuilder();
                    content.Where(s => s.Value == "").ToList()
                        .ForEach(s => emptySections.Append(GetEmptySectionNumber(s.Key)));

                    // value fund check
                    var valueFundCheckResult = new StringBuilder();
                    valuationFund.Where(s => content.TryGetValue(s.Key, out var text) && text.Trim() != "").ToList()
                        .ForEach(s => valueFundCheckResult.Append(s.Value));

                    var row = new TableRow(
                        CreateTableCell($"[{discipline.Code}] {discipline.RussianName}"),
                        CreateTableCell($"{missingSections.Append(emptySections)}"),
                        CreateTableCell($"{valueFundCheckResult}"),
                        CreateTableCell(""),
                        CreateTableCell(""));

                    table.Append(row);
                }
                catch (Exception)
                {
                    continue;
                }
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
            return match.Success ? $"{match.Groups[1].Value}, " : "";
        }

        private static string GetEmptySectionNumber(string s)
        {
            var match = Regex.Match(s, @"^([0-9\.]+).*");
            return match.Success ? $"{match.Groups[1].Value}, " : "";
        }

        private static TableCell CreateTableCell(string text)
            => new(new Paragraph(new Run(new Text(text))));
    }
}
