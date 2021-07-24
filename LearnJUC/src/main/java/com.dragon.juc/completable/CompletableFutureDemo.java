package com.dragon.juc;

import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.util.*;
import java.util.concurrent.*;
import java.util.function.Supplier;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/7/24 12:02
 * @description：
 * @modified By：
 * @version: $
 */
public class CompletableFutureDemo {

    public static void main(String[] args) {
        new Thread(()->{doSearch();}, "Thread_1").start();
        new Thread(()->{doSearch();}, "Thread_2").start();
        new Thread(()->{doSearch();}, "Thread_3").start();
        new Thread(()->{doSearch();}, "Thread_4").start();
        new Thread(()->{doSearch();}, "Thread_5").start();

    }

    private static void doSearch() {
        Map<String, Integer> resultMap = new HashMap<>(16);
        resultMap = completableFutureRun();
        System.out.println(resultMap);
    }

    private static Map<String, Integer> completableFutureRun() {
        ThreadPoolExecutor executor = new ThreadPoolExecutor(3,
                10,
                200, TimeUnit.SECONDS,
                new LinkedBlockingQueue(),
                Executors.defaultThreadFactory(),
                new ThreadPoolExecutor.AbortPolicy());
        Long sMilliSecond = LocalDateTime.now().toInstant(ZoneOffset.of("+8")).toEpochMilli();
        Map<String, Integer> resultMap = new HashMap<>(16);
        CompletableFuture<Map<String, Integer>> futureFlag_1 = CompletableFuture.supplyAsync(new Supplier<Map<String, Integer>>() {
            @Override
            public Map<String, Integer> get() {
                try {
                    TimeUnit.SECONDS.sleep(2);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                resultMap.put("FLAG_1",10);
                return resultMap;
            }
        },executor);

        CompletableFuture<Map<String, Integer>> futureFlag_2 = CompletableFuture.supplyAsync(new Supplier<Map<String, Integer>>() {
            @Override
            public Map<String, Integer> get() {
                try {
                    TimeUnit.SECONDS.sleep(1);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                resultMap.put("FLAG_2",20);
                return resultMap;
            }
        },executor);

        CompletableFuture<Map<String, Integer>> futureFlag_3 = CompletableFuture.supplyAsync(new Supplier<Map<String, Integer>>() {
            @Override
            public Map<String, Integer> get() {
                try {
                    TimeUnit.SECONDS.sleep(3);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                resultMap.put("FLAG_3",30);
                return resultMap;
            }
        },executor);


        CompletableFuture.allOf(futureFlag_1, futureFlag_2, futureFlag_3).join();
        Long eMilliSecond = LocalDateTime.now().toInstant(ZoneOffset.of("+8")).toEpochMilli();
        System.out.println("总耗时：" + (eMilliSecond - sMilliSecond));
        executor.shutdown();
        return resultMap;
    }
}
