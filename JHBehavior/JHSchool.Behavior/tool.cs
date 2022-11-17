using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace JHSchool.Behavior
{
    public static class tool
    {
        public static int SortPeriod(K12.Data.PeriodMappingInfo info1,K12.Data.PeriodMappingInfo info2)
        {
            return info1.Sort.CompareTo(info2.Sort);
        }


        public static Cell SetBordersColor(Cell cell, Color _color)
        {
            Style _style = cell.GetStyle();

            _style.Borders.SetColor(_color);

            cell.SetStyle(_style);
            return cell;
        }

        public static Cell SetBorderType(Cell cell, CellBorderType _type)
        {
            Style _style = cell.GetStyle();

            _style.Borders.SetStyle(_type);

            cell.SetStyle(_style);
            return cell;
        }

        public static Cell SetDiagonalStyle(Cell cell, CellBorderType _type)
        {
            Style _style = cell.GetStyle();

            _style.Borders.DiagonalStyle = _type;

            cell.SetStyle(_style);
            return cell;
        }

        /// <summary>
        /// 文字垂直置中
        /// </summary>
        public static Cell SetHorizontalAlignment(Cell cell, TextAlignmentType _type)
        {
            Style _style = cell.GetStyle();

            _style.HorizontalAlignment = _type;

            cell.SetStyle(_style);
            return cell;
        }

        /// <summary>
        /// 文字水平置中
        /// </summary>
        public static Cell SetVerticalAlignment(Cell cell, TextAlignmentType _type)
        {
            Style _style = cell.GetStyle();

            _style.HorizontalAlignment = _type;

            cell.SetStyle(_style);
            return cell;
        }

        /// <summary>
        /// 字型大小
        /// </summary>
        public static Cell SetFontSize(Cell cell, int fontsize)
        {
            Style _style = cell.GetStyle();

            _style.Font.Size = fontsize;

            cell.SetStyle(_style);
            return cell;
        }

        public static Cell SetShrinkToFit(Cell cell, bool shrinkToFit)
        {
            Style _style = cell.GetStyle();

            _style.ShrinkToFit = shrinkToFit;

            cell.SetStyle(_style);
            return cell;
        }

        public static Range CopyStyle(Range cell, Range range)
        {
            cell.Copy(range);
            cell.CopyStyle(range);
            cell.CopyData(range);
            return cell;
        }
    }
}
