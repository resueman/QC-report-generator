﻿using CurriculumParser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace QCReportGenerator
{
    class QcReportGenerator
    {        
        private Body body;
        private readonly string patternPath;
        private readonly string QCReportPath;
        private readonly List<ProgramRpdsAnalyzer> analysisResults;

        public QcReportGenerator(List<ProgramRpdsAnalyzer> results)
        {
            this.analysisResults = results;
            patternPath = "./pattern.docx";
            QCReportPath = $"./Отчет РПД {results.First().Curriculum.Programme.Code}.docx";
            CreateQCReportDocument();
        }

        public void GenerateReport()
        {
            using var QCReport = WordprocessingDocument.Open(QCReportPath, true);
            body = QCReport.MainDocumentPart.Document.Body;
            foreach (var result in analysisResults.OrderBy(r => r.Course))
            {
                InsertRpdInfo(result);
            }
            InsertAnalytics();
        }

        private void CreateQCReportDocument()
        {
            using var pattern = WordprocessingDocument.Open(patternPath, false);
            using var QCReport = WordprocessingDocument.Create(QCReportPath, WordprocessingDocumentType.Document);
            foreach (var part in pattern.Parts)
            {
                QCReport.AddPart(part.OpenXmlPart, part.RelationshipId);
            }
        }

        private void InsertRpdInfo(ProgramRpdsAnalyzer result)
        {
            var table = body.Descendants<Table>().First();
            var curriculumName = new FileInfo(result.CurriculumPath).Name;

            var tc = CreateTableCellCourseHeader($"{result.Course} курс. Учебный план {curriculumName}");
            table.Append(new TableRow(tc));

            foreach (var (discipline, formMismatchSections, valueFundCheckResult) in result.Results)
            {
                var row = new TableRow(
                    CreateTableCell($"[{discipline.Code}] {discipline.RussianName}"),
                    CreateTableCell($"{formMismatchSections}"),
                    CreateTableCell($"{valueFundCheckResult}"),
                    CreateTableCell(""),
                    CreateTableCell(""));

                table.Append(row);
            }
        }

        private void InsertAnalytics()
        {
            var table = body.Descendants<Table>().First();
            var paragraphs = table.ElementsAfter().Where(e => e is Paragraph).Select(e => e as Paragraph).Skip(2).ToList();
            var info = new List<string>
            {
                $"--- {analysisResults.Sum(r => r.ActualProgramsCount)}/{analysisResults.Sum(r => r.ExpectedProgramsCount)}",
                $"--- {analysisResults.Sum(r => r.IncorrectFormProgramsCount)}",
                $"--- {analysisResults.Sum(r => r.IncorrectValueFundProgramsCount)}"
            };

            paragraphs[0].AppendChild(new Run(new Text(info[0]), new Break()));
            InsertNotAnalyzedRpdInfo(paragraphs[0]);

            for (var i = 1; i < 3; ++i)
            {
                paragraphs[i].AppendChild(new Run(new Text(info[i])));
            }

            InsertProblemsFrequency(paragraphs[3]);
        }

        private void InsertProblemsFrequency(Paragraph paragraph)
        {
            var problemsFrequency = new Dictionary<string, int>();
            foreach (var result in analysisResults)
            {
                foreach (var problem in result.RpdProblemsFrequency)
                {
                    if (!problemsFrequency.ContainsKey(problem.Key))
                    {
                        problemsFrequency.Add(problem.Key, problem.Value);
                        continue;
                    }
                    problemsFrequency[problem.Key] += problem.Value;
                }
            }

            foreach (var problem in problemsFrequency.OrderByDescending(p => p.Value).ToList())
            {
                paragraph.AppendChild(new Run(new Break()));
                paragraph.AppendChild(new Run(new Text($"{problem.Key}  {problem.Value}")));
            }
        }

        private void InsertNotAnalyzedRpdInfo(Paragraph paragraph)
        {
            foreach (var result in analysisResults)
            {
                if (result.IgnoredRpd.Values.All(l => l.Count == 0))
                {
                    return;
                }

                paragraph.AppendChild(CreateRedRun($"Непроанализированные РПД плана {result.CurriculumPath}"));
                foreach (var reason in result.IgnoredRpd.Keys)
                {
                    if (result.IgnoredRpd[reason].Count == 0)
                    {
                        continue;
                    }

                    switch (reason)
                    {
                        case IgnoreReasonType.NotFound:
                            paragraph.AppendChild(CreateRedRun($"РПД, которых не было в {result.RpdFolderPath}:"));
                            break;
                        case IgnoreReasonType.ParsingProblems:
                            paragraph.AppendChild(CreateRedRun("РПД, при парсинге которых возникло исключение:"));
                            break;
                        case IgnoreReasonType.TwoRpdsInFolder:
                            paragraph.AppendChild(CreateRedRun("Несколько РПД для одной дисциплины в папке:"));
                            break;
                    }

                    var counter = 1;
                    foreach (var rpd in result.IgnoredRpd[reason])
                    {
                        paragraph.AppendChild(CreateRedRun($"{counter}. {rpd}"));
                        ++counter;
                    }
                }
            }
        }

        private static TableCell CreateTableCell(string text)
            => new(new Paragraph(new Run(new Text(text))));

        private static Run CreateRedRun(string text) 
            => new(new Text(text), new Break(), new Break())
            {
                RunProperties = new RunProperties { Color = new Color() { Val = "FF0000" } }
            };

        private static TableCell CreateTableCellCourseHeader(string text)
        {
            var js = new Justification { Val = JustificationValues.Center };
            var pPr = new ParagraphProperties() { Justification = js };
            var p = new Paragraph(new Run(new Text(text))) { ParagraphProperties = pPr };
            var tc = new TableCell(p) { TableCellProperties = new TableCellProperties(new GridSpan { Val = 5 }) };
            return tc;
        }
    }
}