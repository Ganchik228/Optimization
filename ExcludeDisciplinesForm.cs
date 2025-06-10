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

        public ExcludeDisciplinesForm(List<Discipline> disciplines, HashSet<string> previouslyExcluded)
        {
            InitializeComponent();
            ExcludedDisciplineNames = new HashSet<string>(previouslyExcluded); // Initialize with previous exclusions

            // Populate checkedListBoxDisciplines
            checkedListBoxDisciplines.Items.Clear();
            var distinctDisciplineNames = disciplines.Select(d => d.Name).Distinct().OrderBy(name => name).ToList();
            foreach (var disciplineName in distinctDisciplineNames)
            {
                checkedListBoxDisciplines.Items.Add(disciplineName, ExcludedDisciplineNames.Contains(disciplineName));
            }
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            ExcludedDisciplineNames.Clear();
            foreach (var item in checkedListBoxDisciplines.CheckedItems)
            {
                if (item != null) // Add null check here
                {
                    ExcludedDisciplineNames.Add(item.ToString()!); // Null-forgiving operator as we've checked for null
                }
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
