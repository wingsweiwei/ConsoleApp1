using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class ExpressionTest
    {
        private class SetterFactory
        {
            public static ISetter<TSource> Create<TSource, TValue>(TSource source, Expression<Func<TSource, TValue>> expression)
            {
                return new Setter<TSource, TValue>(expression);
            }
        }
        private interface ISetter<TSource>
        {
            void Set(TSource source, object value);
        }
        private class Setter<TSource, TValue> : ISetter<TSource>
        {
            private readonly Action<TSource, TValue> action;

            public Setter(Expression<Func<TSource, TValue>> expression)
            {
                this.action = CreateSetter(expression);
            }

            public void Set(TSource source, object value)
            {
                this.action(source, (TValue)value);
            }
            private static Action<TSource, TValue> CreateSetter(Expression<Func<TSource, TValue>> expression)
            {
                if (expression.Body is not MemberExpression memberExpression)
                {
                    throw new ArgumentException("Expression must be a member expression");
                }
                var sourceExpression = Expression.Parameter(typeof(TSource));
                var valueExpression = Expression.Parameter(typeof(TValue));
                var memberAccess = Expression.MakeMemberAccess(sourceExpression, memberExpression.Member);
                var assign = Expression.Assign(memberAccess, valueExpression);
                return Expression.Lambda<Action<TSource, TValue>>(assign, sourceExpression, valueExpression).Compile();
            }
        }
        private class MyClass
        {
            public string Property1 { get; set; }
            public string Property2 { get; set; }
            public string Property3 { get; set; }
        }
    }
}
