using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Optimizations
{
    public class Discipline
    {
        public string Name { get; set; }
        public double MinWorkload { get; set; }
        public double MaxWorkload { get; set; }
        public double SignificanceCoefficient { get; set; }
        public int Semester { get; set; }
        public string UniqueName => $"{Name} (семестр {Semester})";

        public Discipline(string name)
        {
            Name = name;
        }
    }
}
