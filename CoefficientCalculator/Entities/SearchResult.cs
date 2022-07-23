namespace CoefficientCalculator.Entities
{
    public class SearchResult
    {
        public int RowNumber { get; set; }

        public int Wins { get; set; }

        public int Losses { get; set; }

        public int Total { get; set; }

        public decimal Coefficient { get; set; }
    }
}