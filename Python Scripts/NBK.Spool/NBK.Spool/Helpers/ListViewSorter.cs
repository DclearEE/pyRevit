using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;

using System.Threading.Tasks;
using System.Windows.Forms;

namespace NBK.Spool
{
    //
    // Sorts items by column.
    //
    class ListViewSorter : IComparer
    {
        private int col;
        private int doubleCol;
        private HashSet<int> doubleColList = new HashSet<int>();
        private System.Windows.Forms.SortOrder order;

        public ListViewSorter()
        {
            col = 0;
            order = System.Windows.Forms.SortOrder.Ascending;
        }

        public ListViewSorter(int column, System.Windows.Forms.SortOrder order, int doubleColumn)
        {
            col = column;
            doubleCol = doubleColumn;
            doubleColList.Add(doubleCol);
            this.order = order;
        }

        public ListViewSorter(int column, System.Windows.Forms.SortOrder order, HashSet<int> doubleColumns)
        {
            col = column;
            doubleColList = doubleColumns;
            this.order = order;
        }

        public int Compare(object x, object y)
        {
            int returnVal = -1;

            string xString = ((ListViewItem)x).SubItems[col].Text;
            string yString = ((ListViewItem)y).SubItems[col].Text;

            if (doubleColList.Contains(col))
            {
                int test = doubleColList.Single(w => w == col);
                returnVal = Convert.ToDouble(xString).CompareTo(Convert.ToDouble(yString));
            }
            else
            {
                returnVal = String.Compare(xString, yString);
            }

            //
            // Determine whether the sort order is descending.
            // If so, invert the value
            //
            if (order == System.Windows.Forms.SortOrder.Descending)
                returnVal *= -1;

            return returnVal;
        }
    }

    //
    // Specifies how items in a list are sorted.
    //     
    public enum SortOrder
    {
        //
        // Summary:
        //     The items are not sorted.
        None = 0,
        //
        // Summary:
        //     The items are sorted in ascending order.
        Ascending = 1,
        //
        // Summary:
        //     The items are sorted in descending order.
        Descending = 2
    }
}
