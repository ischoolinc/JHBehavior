using System.Collections.Generic;

namespace JHSchool.Behavior.Report
{
    internal class SeatNoComparer : IComparer<string>
    {
        private Dictionary<string, StudentRecord> _mapping;

        public SeatNoComparer(Dictionary<string, StudentRecord> mapping)
        {
            _mapping = mapping;
        }

        #region IComparer<string> 成員

        public int Compare(string x, string y)
        {
            StudentRecord X = _mapping[x];
            StudentRecord Y = _mapping[y];

            int tryX, tryY;
            int intX = (int.TryParse(X.SeatNo, out tryX)) ? tryX : 99999;
            int intY = (int.TryParse(Y.SeatNo, out tryY)) ? tryY : 99999;

            return intX.CompareTo(intY);
        }

        #endregion
    }
}
