using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class Array<T> where T : new()
    {
        readonly SortedDictionary<int, T> data;       

        public Array()
        {
            data = new SortedDictionary<int, T>();   
        }

        public T this[int index]
        {
            get
            {
                T res;
                if (!data.TryGetValue(index, out res))
                {                    
                    res = new T();
                    data.Add(index, res);
                }
                return res;
            }
            set
            {                
                data[index] = value;
            }
        }

        public SortedDictionary<int, T> Data { get { return data; } }
    }    

    public class Matrix<T> where T: new()
    {
        Array<Array<T>> data;

        public Matrix()
        {
            data = new Array<Array<T>>();
        }

        public T this[int row, int col]
        {
            get
            {
                return data[row][col];                
            }
            set
            {
                data[row][col] = value;
            }
        }

        public Array<T> this[int row]
        {
            get { return data[row]; }
            set { data[row] = value; }
        }

        public Array<Array<T>> Data { get { return data; } }
    }
}
