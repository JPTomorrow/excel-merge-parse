using System.Collections.Generic;
using System.Linq;

namespace MoreLinq {
	public static class Ext {
		public static int[] AllIndexs(this string str, params char[] chars) {
			List<int> idx = new List<int>();
			foreach(char c in str) {
				if(chars.Any(x => x.Equals(c)))
					idx.Add(str.IndexOf(c));
			}
			return idx.ToArray();
		}
	}
}