using FISCA.Authentication;
using FISCA.DSAUtil;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Transfer_Excel
{
    [AutoRetryOnWebException()]
    class QueryDiscipline
    {
        public static DSResponse GetDiscipline(DSRequest request)
        {
            return DSAServices.CallService("SmartSchool.Student.Discipline.GetDiscipline", request);
        }
    }
}
