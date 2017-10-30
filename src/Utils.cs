using System;

namespace FluentExcel
{
    /// <summary>
    /// Attribute to mark a method implementation as the defaulted when several signatures are defined
    /// </summary>
    /// <seealso cref="System.Attribute"/>
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
    internal class DefaultImplementationAttribute
        : Attribute
    {
    }

    internal static class Utils
    {
    }
}