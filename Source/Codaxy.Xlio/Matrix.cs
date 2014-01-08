using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    /// <summary>
    /// One dimensional array of objects based on SortedDictionary.
    /// If value does not exists in the dictionary, it's added automatically.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class Array<T> where T : new()
    {
        readonly SortedDictionary<int, T> data;

        /// <summary>
        /// Initializes a new instance of the <see cref="Array{T}"/> class.
        /// </summary>
        public Array()
        {
            data = new SortedDictionary<int, T>();   
        }

        /// <summary>
        /// Gets or sets the cell at the specified index.
        /// </summary>
        /// <value>
        /// The cell.
        /// </value>
        /// <param name="index">The index.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Gets the underlying data.
        /// </summary>
        /// <value>
        /// The data.
        /// </value>
        public SortedDictionary<int, T> Data { get { return data; } }


        /// <summary>
        /// Gets the number of the elements in the array.
        /// </summary>
        /// <value>
        /// The count.
        /// </value>
        public int Count { get { return Data.Count; } }
    }    

    /// <summary>
    /// Two dimensional array.
    /// If value does not exists in the dictionary, it's added automatically.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class Matrix<T> where T: new()
    {
        Array<Array<T>> data;

        /// <summary>
        /// Initializes a new instance of the <see cref="Matrix{T}"/> class.
        /// </summary>
        public Matrix()
        {
            data = new Array<Array<T>>();
        }

        /// <summary>
        /// Gets or sets the cell with the given column and row index.
        /// </summary>
        /// <value>
        /// The cell.
        /// </value>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Gets or sets the row with the specified index.
        /// </summary>
        /// <value>
        /// The row.
        /// </value>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        public Array<T> this[int row]
        {
            get { return data[row]; }
            set { data[row] = value; }
        }

        /// <summary>
        /// Gets the unerlying data structure.
        /// </summary>
        /// <value>
        /// The data.
        /// </value>
        public Array<Array<T>> Data { get { return data; } }
    }
}
