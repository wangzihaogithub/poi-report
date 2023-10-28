package com.github.poireport;

import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.LinkedHashMap;
import java.util.Map;

import static java.util.Locale.ENGLISH;

/**
 * org.springframework.util.ReflectionUtils
 */
public class BeanUtil {

    private static final Map<Class<?>, Field[]> declaredFieldsCache = new LinkedHashMap<Class<?>, Field[]>() {
        @Override
        protected boolean removeEldestEntry(Map.Entry eldest) {
            return size() > 100;
        }
    };

    private static final Field[] EMPTY_FIELD_ARRAY = new Field[0];

    public static Object getFieldValue(String fieldName, Object target) throws IllegalAccessException, NoSuchFieldException {
        try {
            if (target == null) {
                return null;
            }
            PropertyDescriptor descriptor = new PropertyDescriptor(fieldName, target.getClass(),
                    "get" + fieldName.substring(0, 1).toUpperCase(ENGLISH) + fieldName.substring(1),
                    null);
            Method readMethod = descriptor.getReadMethod();
            if (readMethod != null) {
                return readMethod.invoke(target);
            }
        } catch (Exception e) {
            //skip
        }
        Field field = findField(target.getClass(), fieldName, null);
        if (field == null) {
            throw new NoSuchFieldException("field=" + fieldName);
        }
        field.setAccessible(true);
        return field.get(target);
    }

    public static Field findField(Class<?> clazz, String name, Class<?> type) {
        Class<?> searchType = clazz;
        while (Object.class != searchType && searchType != null) {
            Field[] fields = getDeclaredFields(searchType);
            for (Field field : fields) {
                if ((name == null || name.equals(field.getName())) &&
                        (type == null || type.equals(field.getType()))) {
                    return field;
                }
            }
            searchType = searchType.getSuperclass();
        }
        return null;
    }

    private static Field[] getDeclaredFields(Class<?> clazz) {
        Field[] result = declaredFieldsCache.get(clazz);
        if (result == null) {
            try {
                result = clazz.getDeclaredFields();
                declaredFieldsCache.put(clazz, (result.length == 0 ? EMPTY_FIELD_ARRAY : result));
            } catch (Throwable ex) {
                throw new IllegalStateException("Failed to introspect Class [" + clazz.getName() +
                        "] from ClassLoader [" + clazz.getClassLoader() + "]", ex);
            }
        }
        return result;
    }
}
