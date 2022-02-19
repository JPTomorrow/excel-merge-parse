using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Windows.Forms;
using JPMorrow.Tools.Data;
using MoreLinq;
using OfficeOpenXml;

namespace JPMorrow.ExcelMerge
{
	[DataContract]
	public class MergeData {
		[DataMember]
		public List<string[]> Table { get; private set; }

		public MergeData() {
			Table = new List<string[]>();
		}

		public MergeData(List<string[]> row_data) {
			Table = new List<string[]>();
			Table.AddRange(row_data);
		}

		public void AddRow(params string[] row_data) {
			var row = new List<string>();
			row.AddRange(row_data);
			Table.Add(row.ToArray());
		}

		public void DeleteRows(int start_row, int end_row) {
			Table.RemoveRange(start_row, end_row - start_row);
		}

		public List<string[]> FindRows(params string[] search_terms) {

			var ret_rows = new List<string[]>();
			var lower_terms = search_terms.Select(x => x.ToLower()).ToList();
			var add_rows = Table.Where(row => lower_terms.All(term => row.Any(cell => cell.ToLower().Contains(term))));
			ret_rows.AddRange(add_rows);

			return ret_rows;
		}

		public void SaveData(string file_path) {
			try {
				JSON_Serialization.SerializeToFile<MergeData>(this, file_path);
			}
			catch(Exception ex) {
				throw ex;
			}
		}

		public static MergeData LoadData(string file_path) {
			var ret_data = JSON_Serialization.DeserializeFromFile<MergeData>(file_path);
			return ret_data;
		}
	}


	public static class ExMerge {
		public static MergeData MergeExcelFilesIntoTable() {

			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Filter = "Excel Files|*.xlsx;";
			ofd.Title = "Select excel files to merge";
			ofd.Multiselect = true;

			var filenames = new List<string>();

			var result = ofd.ShowDialog();
			if (result == DialogResult.OK)
				filenames.AddRange(ofd.FileNames);

			var packages = new List<ExcelPackage>();

			foreach(var name in filenames)
				packages.Add(new ExcelPackage(new FileInfo(name)));

			var values = new List<string[]>();

			foreach(var p in packages) {
				var ws = p.Workbook.Worksheets;


				foreach(var s in ws) {
					var endidx = s.Dimension.End.Row;

					var row_data = new List<string>();

					for(var ridx = 2; ridx <= endidx; ridx++) {
						var colendidx = s.Dimension.End.Column;

						for(var col = 1; col <= colendidx; col++) {
							var rng = s.Cells[GetExcelColumnName(col).ToString() + ridx.ToString() + ":" + GetExcelColumnName(col).ToString() + ridx.ToString()];
							row_data.Add(new String(Encoding.Convert(Encoding.UTF8, Encoding.Default, rng.Text.Select(c => (byte)c).ToArray()).Select(x => (char)x).ToArray()));
						}
						values.Add(row_data.ToArray());
						row_data.Clear();
					}
				}
			}

			MergeData data = new MergeData(values);
			return data;
		}

		private static string GetExcelColumnName(int columnNumber)
		{
			int dividend = columnNumber;
			string columnName = String.Empty;
			int modulo;

			while (dividend > 0)
			{
				modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
				dividend = (int)((dividend - modulo) / 26);
			}

			return columnName;
		}
	}
}