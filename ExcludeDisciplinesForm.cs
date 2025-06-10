using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Optimizations
{
    public partial class ExcludeDisciplinesForm : Form
    {
        public HashSet<string> ExcludedDisciplineNames { get; private set; }

        public ExcludeDisciplinesForm(List<Discipline> disciplines)
        {
            InitializeComponent();
            ExcludedDisciplineNames = new HashSet<string>();

            var distinctNames = disciplines.Select(d => d.Name).Distinct().OrderBy(name => name).ToList();
            foreach (var name in distinctNames)
            {
                checkedListBoxDisciplines.Items.Add(name, false);
            }
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            foreach (var item in checkedListBoxDisciplines.CheckedItems)
            {
                ExcludedDisciplineNames.Add(item.ToString());
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
