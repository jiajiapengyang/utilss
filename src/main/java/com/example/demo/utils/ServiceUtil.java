package com.example.demo.utils;

import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.annotation.Configuration;

/**
 * @author jiapengyang
 * @summary
 * @Copyright (c) 2017, Lianjia Group All Rights Reserved.
 */
public class ServiceUtil {

    private static ApplicationContext context;

    public ServiceUtil() {
    }

    static void setContext(ApplicationContext applicationContext) {
        context = applicationContext;
    }

    public static <T> T of(Class<T> cls) {
        return context.getBean(cls);
    }

    public static <T> T withName(String beanName, Class<T> cls) {
        return context.getBean(beanName, cls);
    }

    @Configuration
    static class ServicesConfig implements ApplicationContextAware {
        ServicesConfig() {
        }

        @Override
        public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
            ServiceUtil.context = applicationContext;
        }
    }


}
