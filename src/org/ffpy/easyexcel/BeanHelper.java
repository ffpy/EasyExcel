package org.ffpy.easyexcel;

import com.sun.istack.internal.Nullable;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Bean的辅助类
 */
class BeanHelper {
	/** Bean对象类型 */
	private final Class<?> beanClass;
	/** 属性描述符Map（属性名->属性描述符） */
	private final Map<String, PropertyDescriptor> propertyDescriptorMap = new HashMap<>();

	/**
	 * 创建一个BeanHelper实例
	 *
	 * @param beanClass Bean对象类型
	 * @return BeanHelper实例
	 */
	public static BeanHelper of(Class<?> beanClass) {
		return new BeanHelper(beanClass);
	}

	/**
	 * @param beanClass Bean对象类型
	 */
	private BeanHelper(Class<?> beanClass) {
		this.beanClass = beanClass;
		initPropertyDescriptorMap();
	}

	/**
	 * 获取Bean对象类型
	 *
	 * @return Bean对象类型
	 */
	public Class<?> getBeanClass() {
		return beanClass;
	}

	/**
	 * 设置属性的值
	 *
	 * @param bean  Bean对象类型
	 * @param name  属性名
	 * @param value 属性值
	 */
	public void setProperty(Object bean, String name, @Nullable Object value) {
		getPropertyDescriptor(name).setProperty(bean, value);
	}

	/**
	 * 获取属性的值
	 *
	 * @param bean Bean对象类型
	 * @param name 属性名
	 * @param <T>  属性类型
	 * @return 属性的值
	 */
	public <T> T getProperty(Object bean, String name) {
		return getPropertyDescriptor(name).getProperty(bean);
	}

	/**
	 * 获取属性辅助对象
	 *
	 * @param name 属性名
	 * @return 对应的属性辅助对象
	 */
	private PropertyHelper getPropertyDescriptor(String name) {
		PropertyDescriptor property = propertyDescriptorMap.get(name);
		if (property == null)
			throw new IllegalArgumentException("不存在属性" + name);
		return PropertyHelper.of(property);
	}

	/**
	 * 获取Bean的所有属性，按照属性定义顺序排序
	 *
	 * @return 属性辅助对象列表
	 */
	public List<PropertyHelper> getOrderedProperties() {
		List<PropertyHelper> propertyHelperList = new ArrayList<>(propertyDescriptorMap.size());
		for (Field field : beanClass.getDeclaredFields()) {
			if (propertyDescriptorMap.containsKey(field.getName())) {
				propertyHelperList.add(getPropertyDescriptor(field.getName()));
			}
		}
		return propertyHelperList;
	}

	/**
	 * 初始化属性Map
	 */
	private void initPropertyDescriptorMap() {
		try {
			PropertyDescriptor[] properties = Introspector.getBeanInfo(beanClass)
				.getPropertyDescriptors();
			for (PropertyDescriptor property : properties) {
				propertyDescriptorMap.put(property.getName(), property);
			}
		} catch (IntrospectionException e) {
			throw new RuntimeException("读取Bean属性失败", e);
		}
	}
}
