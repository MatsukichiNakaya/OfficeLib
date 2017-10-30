using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace OfficeLib
{
    /// <summary>�l�̌^�ϊ��Ɋւ���g���֐��N���X</summary>
    internal static class ConvertExtention
    {
        /// <summary>
        /// �l�̕ϊ�
        /// </summary>
        /// <typeparam name="TOutput">�ϊ���̌^</typeparam>
        /// <param name="value">�ϊ����Ƃ̒l</param>
        public static TOutput To<TOutput>(this Object value)
        {   
            if(value == null) { return default(TOutput); }

            // �R���o�[�^���쐬���Č^�ϊ����s��
            var converter = TypeDescriptor.GetConverter(typeof(TOutput));
            return converter != null
                    ? (TOutput)converter.ConvertTo(value, typeof(TOutput))
                    : default(TOutput);
        }

        /// <summary>
        /// �^���w�肵�Ă̕ϊ�
        /// </summary>
        /// <param name="value">�ϊ����Ƃ̒l</param>
        /// <param name="type">�ϊ���̌^</param>
        /// <remarks>�߂�l�̃L���X�g���K�v</remarks>
        public static Object To(this Object value, Type type)
        {
            if (value == null)
            {   // Null�̏ꍇ�A�l�^�ł�0, �Q�ƌ^�ł�null��Ԃ��悤�ɂ���
                return type.IsValueType 
                        ? Activator.CreateInstance(type) : null;
            }

            // �R���o�[�^���쐬���Č^�ϊ����s��
            var converter = TypeDescriptor.GetConverter(type);
            return converter != null 
                    ? converter.ConvertTo(value, type) 
                    : (type.IsValueType ? Activator.CreateInstance(type) : null);
        }

        /// <summary>
        /// �l�̕ϊ� �����ŗ�O�������s��
        /// </summary>
        /// <typeparam name="TOutput">�ϊ���̌^</typeparam>
        /// <param name="value">�ϊ����Ƃ̒l</param>
        public static TOutput TryTo<TOutput>(this Object value)
        {
            if (value == null) { return default(TOutput); }

            // �R���o�[�^���쐬���Č^�ϊ����s��
            var converter = TypeDescriptor.GetConverter(typeof(TOutput));
            try
            {
                return converter != null
                        ? (TOutput)converter.ConvertTo(value, typeof(TOutput))
                        : default(TOutput);
            }
            // ��O���̓f�t�H���g�l��Ԃ�
            catch (Exception) { return default(TOutput); }
        }

        /// <summary>
        /// �^���w�肵�Ă̕ϊ� �����ŗ�O�������s��
        /// </summary>
        /// <param name="value">�ϊ����Ƃ̒l</param>
        /// <param name="type">�ϊ���̌^</param>
        /// <remarks>�߂�l�̃L���X�g���K�v</remarks>
        public static Object TryTo(this Object value, Type type)
        {
            // Null�̏ꍇ�A�l�^�ł�0, �Q�ƌ^�ł�null��Ԃ��悤�ɂ���
            Object result = type.IsValueType
                            ? Activator.CreateInstance(type) : null;
            if (value == null) { return result; }
            try
            {
                // �R���o�[�^���쐬���Č^�ϊ����s��
                var converter = TypeDescriptor.GetConverter(type);
                return converter != null
                        ? converter.ConvertTo(value, type) : result;
            }
            // ��O���̓f�t�H���g�l��Ԃ�
            catch (Exception) { return result; }
        }

        /// <summary> �l�̕ϊ�(�ꎟ�z��) </summary>
        /// <typeparam name="TInput">�ϊ����Ƃ̌^</typeparam>
        /// <typeparam name="TOutput">�ϊ���̌^</typeparam>
        /// <param name="values">�ϊ����Ƃ̒l</param>
        public static TOutput[] ConvertAll<TInput, TOutput>(this IEnumerable<TInput> values)
        {
            if (values == null) { return null; }
            return Array.ConvertAll(values.ToArray(), val => val.To<TOutput>());
        }

        /// <summary> �l�̕ϊ�(�񎟔z��) </summary>
        /// <typeparam name="TInput">�ϊ����Ƃ̌^</typeparam>
        /// <typeparam name="TOutput">�ϊ���̌^</typeparam>
        /// <param name="values">�ϊ����Ƃ̒l</param>
        public static TOutput[][] ConvertAll<TInput, TOutput>(this IEnumerable<IEnumerable<TInput>> values)
        {
            if (values == null) { return null; }

            var result = new TOutput[values.Count()][];
            Int32 count = 0;
            foreach (var val in values)
            {
                result[count] = val?.ConvertAll<TInput, TOutput>();
                count++;
            }
            return result;
        }

        /// <summary>
        /// �z���Object�^�֕ϊ�����
        /// </summary>
        /// <typeparam name="T">�ϊ����Ƃ̌^</typeparam>
        /// <param name="src">�ϊ����s���z��</param>
        /// <returns>Object�^�ɕϊ����ꂽ�z��</returns>
        public static Object ToObject<T>(this T[] src)
        {
            return src;
        }

        /// <summary>
        /// Enumerable to array
        /// </summary>
        public static T[] ToArray<T>(this IEnumerable<T> src)
        {
            return Enumerable.ToArray(src);
        }
    }
}


