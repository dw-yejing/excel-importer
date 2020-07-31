package common;

import javax.validation.Constraint;
import javax.validation.Payload;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Constraint(validatedBy = {EnumValidtor.class})

public @interface EnumValidAnnotation {
    String message() default "枚举值不存在";
    Class<?> [] groups() default {};
    Class<? extends Payload>[] payload() default {};
    Class<?>[] target() default {};
}
