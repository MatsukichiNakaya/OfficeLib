using System;

namespace OfficeLib.XLS
{
    /// <summary>
    /// 
    /// </summary>
    public class ComObj : IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        private readonly Object _com;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="com"></param>
        public ComObj(Object com)
        {
            this._com = com;
        }

        ///// <summary>
        ///// 
        ///// </summary>
        //~ComObj() => Dispose();

        /// <summary>
        /// Method
        /// </summary>
        public  ComObj Method(String command)
        {
            return new ComObj(this._com.Method(command));
        }

        /// <summary>
        /// Method Args
        /// </summary>
        public ComObj Method(String command, Object[] args)
        {
            return new ComObj(this._com.Method(command, args));
        }

        /// <summary>
        /// Getproperty
        /// </summary>
        public ComObj GetProperty(String command)
        {
            return new ComObj(this._com.GetProperty(command));
        }

        /// <summary>
        /// Getproperty  Args
        /// </summary>
        public ComObj GetProperty(String command, Object[] args)
        {
            return new ComObj(this._com.GetProperty(command, args));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public Int32 ToInt()
        {
            return this._com.To<Int32>();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            OfficeCore.ReleaseObject(this._com);
        }
    }
}
