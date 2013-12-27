using System;
using System.Collections.Generic;
using System.Text;

namespace JHSchool.Behavior.Legacy
{
    public interface ICellInfo<T>
    {
        T OriginValue { get;}
        void SetValue(T value);
        bool IsDirty { get;}
    } 
}
