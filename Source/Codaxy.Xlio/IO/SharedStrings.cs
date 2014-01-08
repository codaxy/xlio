using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.IO
{
    class SharedStrings : IEnumerable<string>
    {
        Dictionary<String, int> lookup;
        List<String> data;

        public SharedStrings()
        {
            lookup = new Dictionary<string, int>();
            data = new List<string>();
        }

        public int Add(String s)
        {
            int index = data.Count;
            data.Add(s);
            lookup[s] = index;
            return index;
        }

        public int Get(String s)
        {
            int index;
            if (lookup.TryGetValue(s, out index))
                return index;
            return Add(s);
        }

        public String this[int index]
        {
            get
            {
                if (index >= 0 && index < data.Count)
                    return data[index];
                return null;
            }
        }


        public IEnumerator<string> GetEnumerator()
        {
            return data.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return data.GetEnumerator();
        }

        public int Count { get { return data.Count; } }
    }
}
