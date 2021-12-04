package com.dragon.config;

import com.dragon.common.utils.JasyptUtil;
import org.redisson.Redisson;
import org.redisson.config.Config;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.data.redis.connection.lettuce.LettuceConnectionFactory;
import org.springframework.data.redis.core.RedisTemplate;
import org.springframework.data.redis.serializer.GenericJackson2JsonRedisSerializer;
import org.springframework.data.redis.serializer.StringRedisSerializer;

import java.io.Serializable;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021-11-14 22:02
 * @description：
 * @modified By：
 * @version: $
 */
@Configuration
public class RedisConfig {

    @Value("${jasypt.encryptor.password}")
    private String password;

    @Bean
    public RedisTemplate<String, Serializable> redisTemplate(LettuceConnectionFactory connectionFactory) {
        RedisTemplate<String, Serializable> redisTemplate = new RedisTemplate<>();

        redisTemplate.setKeySerializer(new StringRedisSerializer());
        redisTemplate.setValueSerializer(new GenericJackson2JsonRedisSerializer());
        redisTemplate.setConnectionFactory(connectionFactory);

        return redisTemplate;
    }

    @Bean
    public Redisson redisson() {
        Config config = new Config();
        String redisPassword = JasyptUtil.decyptPwd(password, "r9B8FJXprGJ/0yHlpt12/5JEfCfCuaNwaNX+PzAKpFA=");
        config.useSingleServer().setAddress("redis://81.69.249.173:6380").setDatabase(0).setPassword(redisPassword);

        return (Redisson) Redisson.create(config);
    }
}
