package indi.liht.stat;

import indi.liht.stat.core.IStat;
import indi.liht.stat.core.StatMovie;

import java.util.concurrent.TimeUnit;

/**
 * Usage:
 * 程序入口
 * @author lihongtao ibraxwell@sina.com
 * on 2018/10/28
 **/
public class Main {

    /**
     * 程序入口
     * @param args 入口参数
     */
    public static void main(String[] args) {
        // 开始统计
        IStat stat = new StatMovie();
        System.out.println("やほー、统计开始……");
        stat.stat();
        try {
            System.out.println("よっしゃ、mission completed！5秒后退出程序……");
            TimeUnit.SECONDS.sleep(5);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

}
