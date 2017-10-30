using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq.Expressions;
using System.Reflection;

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
        internal static string GetColumnTitle(LambdaExpression expr, string separator = " ")
        {
            var stack = new Stack<string>();

            MemberExpression me;
            switch (expr.Body.NodeType)
            {
                case ExpressionType.Convert:
                case ExpressionType.ConvertChecked:
                    var ue = expr.Body as UnaryExpression;
                    me = ((ue != null) ? ue.Operand : null) as MemberExpression;
                    break;

                default:
                    me = expr.Body as MemberExpression;
                    break;
            }

            while (me != null)
            {
                stack.Push(me.Member.GetCustomAttribute<DisplayAttribute>(true)?.Name ?? me.Member.Name);
                me = me.Expression as MemberExpression;
            }

            return string.Join(separator, stack.ToArray());
        }
    }
}