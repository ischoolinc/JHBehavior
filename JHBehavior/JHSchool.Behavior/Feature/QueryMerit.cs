using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.DSAUtil;

namespace JHSchool.Behavior.Feature
{
    [FISCA.Authentication.AutoRetryOnWebException()]
    public class QueryMerit
    {
        public static DSResponse GetMeritStatistic(DSRequest request)
        {
            return FISCA.Authentication.DSAServices.CallService("SmartSchool.Student.Discipline.GetMeritStatistic", request);
        }

        public static DSResponse GetMeritIgnoreDemerit(DSRequest request)
        {
            return FISCA.Authentication.DSAServices.CallService("SmartSchool.Student.Discipline.GetMeritIgnoreDemerit", request);
        }

        public static DSResponse GetMeritIgnoreUnclearedDemerit(DSRequest request)
        {
            return FISCA.Authentication.DSAServices.CallService("SmartSchool.Student.Discipline.GetMeritIgnoreUnclearedDemerit", request);
        }


        public static DSResponse GetDiscipline(DSRequest request)
        {
            return FISCA.Authentication.DSAServices.CallService("SmartSchool.Student.Discipline.GetDiscipline", request);
        }
    }
}
