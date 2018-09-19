package org.ffpy.easyexcel;

import com.sun.istack.internal.Nullable;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * Bean的工具类
 */
class BeanUtils {

	/**
	 * 设置属性的值
	 *
	 * @param bean  bean实例
	 * @param name  属性名
	 * @param value 属性值
	 */
	public static void setProperty(Object bean, String name, @Nullable Object value) {
		try {
			PropertyDescriptor descriptor = getPropertyDescriptor(bean.getClass(), name);
			if (descriptor == null)
				throw new IllegalArgumentException("不存在属性" + name);
			Method method = descriptor.getWriteMethod();
			if (method == null)
				throw new IllegalArgumentException(name + "没有setter方法");
			method.invoke(bean, value);
		} catch (IllegalAccessException | InvocationTargetException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 获取属性的值
	 *
	 * @param bean bean实例
	 * @param name 属性名
	 * @param <T>  属性类型
	 * @return 属性的值
	 */
	public static <T> T getProperty(Object bean, String name) {
		try {
			PropertyDescriptor descriptor = getPropertyDescriptor(bean.getClass(), name);
			if (descriptor == null)
				throw new IllegalArgumentException("不存在属性" + name);
			Method method = descriptor.getReadMethod();
			if (method == null)
				throw new IllegalArgumentException(name + "没有getter方法");
			//noinspection unchecked
			return (T) method.invoke(bean);
		} catch (IllegalAccessException | InvocationTargetException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 获取属性的描述符
	 *
	 * @param beanClass bean的类型
	 * @param name      属性名
	 * @return 属性的描述符，如果没有找到该属性，则返回null
	 */
	private static PropertyDescriptor getPropertyDescriptor(Class<?> beanClass, String name) {
		try {
			PropertyDescriptor[] descriptors = Introspector.getBeanInfo(beanClass).getPropertyDescriptors();
			for (PropertyDescriptor descriptor : descriptors) {
				if (descriptor.getName().equals(name)) {
					return descriptor;
				}
			}
		} catch (IntrospectionException e) {
			throw new RuntimeException(e);
		}
		return null;
	}

	private BeanUtils() {
	}
}
