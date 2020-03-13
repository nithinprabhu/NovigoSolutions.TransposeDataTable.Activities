using System.Activities;
using System.ComponentModel;

namespace NovigoSolutions.TransposeDataTable.Activities
{
    [Category("Novigo Solutions.DataTable")]
    [DisplayName("Transpose DataTable")]
    public class TransposeDataTable : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("The DataTable to transpose.")]
        [DisplayName("DataTable")]
        public InArgument<System.Data.DataTable> InputDataTable { get; set; }

        [Category("Options")]
        [Description("Column name to be set for rows that begin with empty cell. By default set to 'Column' and appended by incrementing number starting from 1.")]
        [DisplayName("Alternate Column Name")]
        public InArgument<string> AlternateColumnName { get; set; }

        [Category("Output")]
        [Description("The resulting transposed DataTable.")]
        [DisplayName("DataTable")]
        public OutArgument<System.Data.DataTable> OutputDataTable { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            OutputDataTable.Set(context,GenerateTransposedDataTable(this.InputDataTable.Get(context), this.AlternateColumnName.Get(context)));
        }

        private static System.Data.DataTable GenerateTransposedDataTable(System.Data.DataTable inputDataTable,string altColumnName)
        {
            System.Data.DataTable outputDataTable = new System.Data.DataTable();
            outputDataTable.Columns.Add(inputDataTable.Columns[0].ColumnName.ToString());

            int altColumnCount = 1;
            foreach (System.Data.DataRow inRow in inputDataTable.Rows)
            {
                string newColName = inRow[0].ToString();
                if (string.IsNullOrEmpty(newColName))
                {
                    altColumnName=string.IsNullOrWhiteSpace(altColumnName) ? "Column" : altColumnName;
                    outputDataTable.Columns.Add(string.Concat(altColumnName,altColumnCount));
                    altColumnCount++;
                } else if (outputDataTable.Columns.Contains(newColName))
                {
                    outputDataTable.Columns.Add(string.Concat(newColName, altColumnCount));
                    altColumnCount++;
                } else
                {
                    outputDataTable.Columns.Add(newColName);
                }
            }
     
            for (int rowCount = 1; rowCount <= inputDataTable.Columns.Count - 1; rowCount++)
            {
                System.Data.DataRow newRow = outputDataTable.NewRow();

                newRow[0] = inputDataTable.Columns[rowCount].ColumnName.ToString();
                for (int columnCount = 0; columnCount <= inputDataTable.Rows.Count - 1; columnCount++)
                {
                    string colValue = inputDataTable.Rows[columnCount][rowCount].ToString();
                    newRow[columnCount + 1] = colValue;
                }
                outputDataTable.Rows.Add(newRow);
            }
            return outputDataTable;
        }
    }
}