package config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.validation.beanvalidation.CustomValidatorBean;

@Configuration
public class ValidatorConfig {
    @Bean
    public CustomValidatorBean customValidatorBean(){
        return new CustomValidatorBean();
    }
}
