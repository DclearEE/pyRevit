using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NBK.Spool.Forms
{
    public partial class SpoolForm : Form
    {
        private List<SpoolNumber> m_spoolNumbers;
        private string m_newSpoolNumber;
        private int sortColumn = -1;

        public string NewSpoolNumber;


        public SpoolForm(List<SpoolNumber> spoolNumbers)
        {
            this.m_spoolNumbers = spoolNumbers;
            InitializeComponent();
        }

        private void SpoolForm_Load(object sender, EventArgs e)
        {
            PopulateListView();
            //
            // To esure the dialog renders with the proper font, get all the buttons and labels and set their text rendering to false
            //
            MicrodeskHelpers.DisableTextRendering(this);
        }

        private void PopulateListView()
        {
            listView.Items.Clear();
            this.listView.ListViewItemSorter = null;

            foreach (SpoolNumber spool in m_spoolNumbers)
            {
                ListViewItem newItem = new ListViewItem();

                newItem.Text = spool.FullSpoolNo;
                newItem.Tag = spool;

                newItem.SubItems.Add(spool.Abbreviation1);
                newItem.SubItems.Add(spool.Abbreviation2);
                newItem.SubItems.Add(spool.Number.ToString());

                listView.Items.Add(newItem);
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.NewSpoolNumber = label_SpoolNo.Text;
        }

        private void textBoxAbbr1_TextChanged(object sender, EventArgs e)
        {
            NewSpoolNo();
        }

        private void textBoxAbbr2_TextChanged(object sender, EventArgs e)
        {
            NewSpoolNo();
        }

        private void NewSpoolNo()
        {
            label_SpoolNo.Text = textBoxAbbr1.Text + "-" + textBoxAbbr2.Text + "-" + textBoxNumber.Text;
        }

        private void textBoxNumber_TextChanged(object sender, EventArgs e)
        {
            NewSpoolNo();
        }

        private void listView_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListView.SelectedListViewItemCollection selectedItems = this.listView.SelectedItems;

            if (selectedItems.Count == 0) return;

            textBoxAbbr1.Text = selectedItems[0].SubItems[1].Text;
            textBoxAbbr2.Text = selectedItems[0].SubItems[2].Text;

            List<SpoolNumber> localSpools = this.m_spoolNumbers
                .Where(x => x.Abbreviation1 == textBoxAbbr1.Text)
                .Where(x => x.Abbreviation2 == textBoxAbbr2.Text)
                .OrderBy(x => x.Number)
                .ToList();

            textBoxNumber.Text = (localSpools.Last().Number + 1).ToString();
        }

        private void listView_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
        {
            try
            {
                //
                // Determine whether the column is the same as the last column clicked.
                //
                if (e.Column != sortColumn)
                {
                    //
                    // Set the sort column to the new column.
                    //
                    sortColumn = e.Column;

                    //
                    // Set the sort order to ascending by default.
                    //
                    listView.Sorting = System.Windows.Forms.SortOrder.Ascending;
                }
                else
                {
                    //
                    // Determine what the last sort order was and change it.
                    //
                    if (listView.Sorting == System.Windows.Forms.SortOrder.Ascending)
                        listView.Sorting = System.Windows.Forms.SortOrder.Descending;
                    else
                        listView.Sorting = System.Windows.Forms.SortOrder.Ascending;
                }

                //
                // Set the ListViewItemSorter property to a new ListViewItemComparer object.
                //
                listView.BeginUpdate();
                this.listView.ListViewItemSorter = new ListViewSorter(e.Column, listView.Sorting, 1);
                listView.EndUpdate();
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<SpoolNumber> localSpools = this.m_spoolNumbers
                .Where(x => x.Abbreviation1 == textBoxAbbr1.Text)
                .Where(x => x.Abbreviation2 == textBoxAbbr2.Text)
                .OrderBy(x => x.Number)
                .ToList();

            textBoxNumber.Text = (localSpools.Last().Number + 1).ToString();
        }
    }
}
