namespace QCReportGenerator
{
    /// <summary>
    /// Причины, по которым рабочая программа дисциплины не была проанализированна => не занесена в таблицу
    /// </summary>
    public enum IgnoreReasonType
    {
        ParsingProblems,
        NotFound,
        TwoRpdsInFolder
    }
}