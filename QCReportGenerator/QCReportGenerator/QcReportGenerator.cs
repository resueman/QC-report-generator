using CurriculumParser;
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
        private readonly ProgramRpdsAnalyzer programRpdsAnalyzer;

        public QcReportGenerator(ProgramRpdsAnalyzer programRpdsAnalyzer)
        {
            this.programRpdsAnalyzer = programRpdsAnalyzer;
            patternPath = "./pattern.docx";
            QCReportPath = $"./Отчет РПД {new FileInfo(programRpdsAnalyzer.CurriculumPath).Name}";
            CreateQCReportDocument();
        }

        public void GenerateReport()
        {
            using var QCReport = WordprocessingDocument.Open(QCReportPath, true);
            body = QCReport.MainDocumentPart.Document.Body;
            foreach (var (discipline, formMismatchSections, valueFundCheckResult) in programRpdsAnalyzer.Results)
            {
                InsertRpdInfo(discipline, formMismatchSections, valueFundCheckResult);
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

        private void InsertRpdInfo(Discipline discipline, string formMismatchSections, string valueFundCheckResult)
        {
            var table = body.Descendants<Table>().First();
            var row = new TableRow(
                CreateTableCell($"[{discipline.Code}] {discipline.RussianName}"),
                CreateTableCell($"{formMismatchSections}"),
                CreateTableCell($"{valueFundCheckResult}"),
                CreateTableCell(""),
                CreateTableCell(""));

            table.Append(row);
        }

        private void InsertAnalytics()
        {
            var table = body.Descendants<Table>().First();
            var paragraphs = table.ElementsAfter().Where(e => e is Paragraph).Skip(2).ToList();
            var info = new List<string>
            {
                $"--- {programRpdsAnalyzer.ActualProgramsCount}/{programRpdsAnalyzer.ExpectedProgramsCount}",
                $"--- {programRpdsAnalyzer.IncorrectFormProgramsCount}",
                $"--- {programRpdsAnalyzer.IncorrectValueFundProgramsCount}"
            };

            paragraphs[0].AppendChild(new Run(new Text(info[0])));
            paragraphs[0].AppendChild(new Run(new Break()));
            InsertNotAnalyzedRpdInfo(paragraphs[0] as Paragraph);

            for (var i = 1; i < 3; ++i)
            {
                paragraphs[i].AppendChild(new Run(new Text(info[i])));
            }

            foreach (var problem in programRpdsAnalyzer.RpdProblemsFrequency.OrderByDescending(p => p.Value).ToList())
            {
                paragraphs[3].AppendChild(new Run(new Break()));
                paragraphs[3].AppendChild(new Run(new Text($"{problem.Key}  {problem.Value}")));
            }
        }

        private void InsertNotAnalyzedRpdInfo(Paragraph paragraph)
        {
            foreach (var reason in programRpdsAnalyzer.IgnoredRpd.Keys)
            {
                if (programRpdsAnalyzer.IgnoredRpd[reason].Count == 0)
                {
                    return;
                }

                switch (reason)
                {
                    case IgnoreReasonType.NotFound:
                        paragraph.AppendChild(CreateRedRun("РПД, которых не было в указанной папке:"));
                        break;
                    case IgnoreReasonType.ParsingProblems:
                        paragraph.AppendChild(CreateRedRun("РПД, при парсинге которых возникло исключение:"));
                        break;
                }

                var counter = 1;
                foreach (var rpd in programRpdsAnalyzer.IgnoredRpd[reason])
                {
                    paragraph.AppendChild(CreateRedRun($"{counter}. {rpd}"));
                    ++counter;
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
    }
}
