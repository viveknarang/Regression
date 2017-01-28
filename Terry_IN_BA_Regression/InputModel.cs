using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Terry_IN_BA_Regression
{
    public class InputModel
    {
        public HashSet<String> columns;
        public int totalItems;
        public string[,] array;
        public string[,] arrayWithObservationNumbers;
        public string cellnames;
        public Dictionary<string, int>  observationsMap = new Dictionary<string, int>();
    }
}
