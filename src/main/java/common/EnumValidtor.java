package common;

import javax.validation.ConstraintValidator;
import javax.validation.ConstraintValidatorContext;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Objects;

public class EnumValidtor implements ConstraintValidator<EnumValidAnnotation, String> {
    private Class<?>[] clazzs;

    @Override
    public void initialize(EnumValidAnnotation constrainAnnotation){
        clazzs = constrainAnnotation.target();
    }

    @Override
    public boolean isValid(String value, ConstraintValidatorContext constraintValidatorContext){
        if(clazzs.length > 0){
            for(Class clazz : clazzs ){
                try {
                    if(clazz.isEnum()){
                        Object[] objects = clazz.getEnumConstants();
                        Method method = clazz.getMethod("getInfo");
                        for(Object object : objects){
                            Object code = method.invoke(object, null);
                            if(Objects.equals(value, code.toString())){
                                return true;
                            }
                        }
                    }
                } catch (NoSuchMethodException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
        return false;
    }
}
