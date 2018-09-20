package org.ffpy.easyexcel;

import com.sun.istack.internal.Nullable;

import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * 属性辅助类
 */
class PropertyHelper {
	private PropertyDescriptor propertyDescriptor;

	/**
	 * 创建一个PropertyHelper对象
	 *
	 * @param propertyDescriptor 属性描述符
	 * @return PropertyHelper对象
	 */
	public static PropertyHelper of(PropertyDescriptor propertyDescriptor) {
		return new PropertyHelper(propertyDescriptor);
	}

	/**
	 * @param propertyDescriptor 属性描述符
	 */
	private PropertyHelper(PropertyDescriptor propertyDescriptor) {
		this.propertyDescriptor = propertyDescriptor;
	}

	/**
	 * 获取属性描述符
	 *
	 * @return 属性描述符
	 */
	public PropertyDescriptor getPropertyDescriptor() {
		return propertyDescriptor;
	}

	/**
	 * 获取属性的值
	 *
	 * @param bean Bean对象类型
	 * @param name 属性名
	 * @param <T>  属性类型
	 * @return 属性的值
	 */
	public <T> T getProperty(Object bean) {
		try {
			Method method = propertyDescriptor.getReadMethod();
			if (method == null)
				throw new IllegalArgumentException(getName() + "没有getter方法");
			//noinspection unchecked
			return (T) method.invoke(bean);
		} catch (IllegalAccessException | InvocationTargetException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 设置属性的值
	 *
	 * @param bean  Bean对象类型
	 * @param name  属性名
	 * @param value 属性值
	 */
	public void setProperty(Object bean, @Nullable Object value) {
		try {
			Method method = propertyDescriptor.getWriteMethod();
			if (method == null)
				throw new IllegalArgumentException(getName() + "没有setter方法");
			method.invoke(bean, value);
		} catch (IllegalAccessException | InvocationTargetException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 获取属性类型
	 *
	 * @return 属性类型
	 */
	public Class<?> getPropertyType() {
		return propertyDescriptor.getPropertyType();
	}

	/**
	 * 获取属性名
	 *
	 * @return 属性名
	 */
	public String getName() {
		return propertyDescriptor.getName();
	}
}
