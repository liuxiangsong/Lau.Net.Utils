using AutoMapper;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils
{
    public static class AutoMapperUtil
    {
        /// <summary>
        /// 类型映射
        /// </summary>
        /// <typeparam name="TDestination">目标对象类型</typeparam>
        /// <param name="obj">要映射的对象</param>
        /// <returns></returns>
        public static TDestination MapTo<TDestination>(object source) where TDestination : class
        {
            if (source == null)
            {
                return default(TDestination);
            }
            var config = new MapperConfiguration(cfg => cfg.CreateMap(source.GetType(), typeof(TDestination)));
            var mapper = config.CreateMapper();
            return mapper.Map<TDestination>(source);
        }

        /// <summary>
        /// 类型映射
        /// </summary>
        /// <typeparam name="TDestination">目标对象类型</typeparam>
        /// <param name="source">要映射的对象</param>
        /// <param name="configure"></param>
        /// <returns></returns>
        public static TDestination MapTo<TDestination>(object source, Action<IMapperConfigurationExpression> configure) where TDestination : class
        {
            if (source == null)
            {
                return default(TDestination);
            }
            var config = new MapperConfiguration(configure);
            var mapper = config.CreateMapper();
            return mapper.Map<TDestination>(source);
        }

        /// <summary>
        /// 合并对象，TSource的字段覆盖TDestination中相同的字段,TDestination中独有的字段保持
        /// </summary>
        /// <typeparam name="TSource">数据源类型</typeparam>
        /// <typeparam name="TDestination">目标对象类型</typeparam>
        /// <param name="source">数据源</param>
        /// <param name="destination">目标对象</param>
        /// <returns></returns>
        public static TDestination MergeTo<TSource, TDestination>(TSource source, TDestination destination)
            where TSource : class
            where TDestination : class
        {
            if (source == null)
            {
                return destination;
            }
            var config = new MapperConfiguration(cfg => cfg.CreateMap<TSource, TDestination>());
            var mapper = config.CreateMapper();
            return mapper.Map<TSource, TDestination>(source, destination);
        }

        /// <summary>
        /// 集合列表类型映射
        /// </summary>
        /// <typeparam name="TDestination">目标对象类型</typeparam>
        /// <param name="source">数据源</param>
        /// <returns></returns>
        public static List<TDestination> MapTo<TDestination>(IEnumerable source) where TDestination : class
        {
            if (source == null)
            {
                return null;
            }

            var config = new MapperConfiguration(cfg => {
                cfg.CreateMap(source.GetType().GetGenericArguments()[0], typeof(TDestination));
            });
            var mapper = config.CreateMapper();
            return mapper.Map<List<TDestination>>(source);
        }
    }

}
