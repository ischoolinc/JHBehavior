using System;
using System.Collections.Generic;
using System.Text;

namespace JHSchool.Legacy
{
    public class Discipline
    {
        private string _occurDate;

        public string OccurDate
        {
            get { return _occurDate; }
            set { _occurDate = value; }
        }
        private string _a, _b, _c;

        public string C
        {
            get { return _c; }
            set { _c = value; }
        }

        public string B
        {
            get { return _b; }
            set { _b = value; }
        }

        public string A
        {
            get { return _a; }
            set { _a = value; }
        }
        private string _reason;

        public string Reason
        {
            get { return _reason; }
            set { _reason = value; }
        }

        private string _schoolYear;

        public string SchoolYear
        {
            get { return _schoolYear; }
            set { _schoolYear = value; }
        }
        private string _semester;

        public string Semester
        {
            get { return _semester; }
            set { _semester = value; }
        }
        private string _gradeYear;

        public string GradeYear
        {
            get { return _gradeYear; }
            set { _gradeYear = value; }
        }
        private string _id;

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }
        public Discipline()
        { }

        public override string ToString()
        {
            return _id;
        }

        private bool _cleared;

        public bool Cleared
        {
            get { return _cleared; }
            set { _cleared = value; }
        }
        private string _clearDate;

        public string ClearDate
        {
            get { return _clearDate; }
            set { _clearDate = value; }
        }
        private string _clearReason;

        public string ClearReason
        {
            get { return _clearReason; }
            set { _clearReason = value; }
        }

        private bool _isAsshole;

        public bool IsAsshole
        {
            get { return _isAsshole; }
            set { _isAsshole = value; }
        }
    }
}
