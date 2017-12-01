using System;
using System.Reflection;

namespace OfficeLib
{
    /// <summary>Reflection InvokeMember extention class</summary>
    internal static class InvokeExtention
    {
        #region --- InvokeMethod ---
        /// <summary>
        /// Invoke Method extention
        /// </summary>
        public static Object Method(this Object src, String command)
        {
            return Method(src, command, null, src, null);
        }

        /// <summary>
        /// Invoke Method extention
        /// </summary>
        public static Object Method(this Object src, String command, Object[] args)
        {
            return Method(src, command, null, src, args);
        }

        /// <summary>
        /// Invoke Method extention
        /// </summary>
        public static Object Method(this Object src, String command, Object target, Object[] args)
        {
            return Method(src, command, null, target, args);
        }

        /// <summary>
        /// Invoke Method extention
        /// </summary>
        public static Object Method(this Object src, String command, Binder binder, Object target, Object[] args)
        {
            if (src == null) { return null; }
            return src.GetType().InvokeMember(command, BindingFlags.InvokeMethod, binder, target, args);
        }
        #endregion

        #region --- Getproperty ---
        /// <summary>
        /// Invoke Getproperty extention
        /// </summary>
        public static Object GetProperty(this Object src, String command)
        {
            return GetProperty(src, command, null, src, null);
        }

        /// <summary>
        /// Invoke Getproperty extention
        /// </summary>
        public static Object GetProperty(this Object src, String command, Object[] args)
        {
            return GetProperty(src, command, null, src, args);
        }

        /// <summary>
        /// Invoke Getproperty extention
        /// </summary>
        public static Object GetProperty(this Object src, String command, Object target, Object[] args)
        {
            return GetProperty(src, command, null, target, args);
        }

        /// <summary>
        /// Invoke Getproperty extention
        /// </summary>
        public static Object GetProperty(this Object src, String command, Binder binder, Object target, Object[] args)
        {
            if(src == null) { return null; }
            return src.GetType().InvokeMember(command, BindingFlags.GetProperty, binder, target, args);
        }
        #endregion

        #region --- SetProperty ---
        ///// <summary>
        ///// Invoke Setproperty extention
        ///// </summary>
        //public static Object SetProperty(this Object src, String command)
        //{
        //    return SetProperty(src, command, null, src, null);
        //}

        /// <summary>
        /// Invoke Setproperty extention
        /// </summary>
        public static Object SetProperty(this Object src, String command, Object[] args)
        {
            return SetProperty(src, command, null, src, args);
        }

        /// <summary>
        /// Invoke Setproperty extention
        /// </summary>
        public static Object SetProperty(this Object src, String command, Object target, Object[] args)
        {
            return SetProperty(src, command, null, target, args);
        }

        /// <summary>
        /// Invoke Setproperty extention
        /// </summary>
        public static Object SetProperty(this Object src, String command, Binder binder, Object target, Object[] args)
        {
            if (src == null) { return null; }
            return src.GetType().InvokeMember(command, BindingFlags.SetProperty, binder, target, args);
        }
        #endregion
    }
}
