using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace QCReportGenerator
{
    /// <summary>
    /// Генератор отчета комиссии контроля качества
    /// </summary>
    class QcReportGenerator
    {        
        private Body body;
        private readonly string patternPath;
        private readonly string QCReportPath;
        private readonly List<ProgramRpdsAnalyzer> analysisResults;

        public QcReportGenerator(List<ProgramRpdsAnalyzer> results)
        {
            analysisResults = results;
            patternPath = "./pattern.docx";
            QCReportPath = $"./Отчет РПД {results.First().Curriculum.Programme.Code}.docx";
            CreateQCReportDocument();
        }

        /// <summary>
        /// Создать отчет комиссии контроля качества
        /// </summary>
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

        /// <summary>
        /// Создать документ отчета по шаблону
        /// </summary>
        private void CreateQCReportDocument()
        {
            using var pattern = WordprocessingDocument.Open(patternPath, false);
            using var QCReport = WordprocessingDocument.Create(QCReportPath, WordprocessingDocumentType.Document);
            foreach (var part in pattern.Parts)
            {
                QCReport.AddPart(part.OpenXmlPart, part.RelationshipId);
            }
        }

        /// <summary>
        /// Вставляет в документ строку с наименованием дисциплины и ячейками для вписания результатов 
        /// проверки рабочей программы данной дисциплины
        /// </summary>
        /// <param name="result">Результат проверки рабочих программ конкретного направления обучения курса бакалавриата будущего года</param>
        private void InsertRpdInfo(ProgramRpdsAnalyzer result)
        {
            var table = body.Descendants<Table>().First();
            var curriculumName = Path.GetFileNameWithoutExtension(result.CurriculumPath);

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

        /// <summary>
        /// Вставляет в документ информацию для '2. Аналитические выводы'
        /// </summary>
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

        /// <summary>
        /// Вставляет в документ информацию о количестве ошибок для каждой секции РПД
        /// </summary>
        /// <param name="paragraph">Параграф, в конец которого добавляется информация</param>
        private void InsertProblemsFrequency(Paragraph paragraph)
        {
            var problemsFrequency = new Dictionary<string, int>();
            foreach (var result in analysisResults)
            {
                foreach (var problem in result.RpdSectionProblemsFrequency)
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

        /// <summary>
        /// Вставляет в документ информацию о РПД, которые не были проанализированны => не включены в таблицу
        /// </summary>
        /// <param name="paragraph">Параграф, в конец которого добавляется информация</param>
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
                            paragraph.AppendChild(CreateRedRun($"РПД, при разборе которых возникло исключение (план {result.CurriculumPath}):"));;
                            break;
                        case IgnoreReasonType.TwoRpdsInFolder:
                            paragraph.AppendChild(CreateRedRun($"Несколько РПД для одной дисциплины в папке (план {result.CurriculumPath}):"));
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

        /// <summary>
        /// Создает ячейку таблицы, заданной в шаблоне
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static TableCell CreateTableCell(string text)
            => new(new Paragraph(new Run(new Text(text))));

        /// <summary>
        /// Создает пробег с текстом красного цвета для вставки информации о необработанных РПД
        /// </summary>
        private static Run CreateRedRun(string text) 
            => new(new Text(text), new Break(), new Break())
            {
                RunProperties = new RunProperties { Color = new Color() { Val = "FF0000" } }
            };

        /// <summary>
        /// Создает строку в таблице с одной ячейкой, которая описывает, РПД какого курса бакалавриата анализируются
        /// </summary>
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
