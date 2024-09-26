package com.ransibi.config;

import ch.qos.logback.classic.Level;
import ch.qos.logback.classic.LoggerContext;
import ch.qos.logback.classic.spi.LoggerContextListener;
import ch.qos.logback.core.Context;
import ch.qos.logback.core.spi.ContextAwareBase;
import ch.qos.logback.core.spi.LifeCycle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

/**
 * @description: 获取日志最大保留时间
 * @author: yjd-闫九鼎
 * @create: 2022-06-17 11:46
 * @Version: 1.0
 */
public class LogHistoryListener extends ContextAwareBase
        implements LoggerContextListener, LifeCycle {
    private static final Logger logger = LoggerFactory.getLogger(LogHistoryListener.class);

    private boolean started = false;

    @Override
    public boolean isResetResistant() {
        return false;
    }

    @Override
    public void onStart(LoggerContext context) {

    }

    @Override
    public void onReset(LoggerContext context) {

    }

    @Override
    public void onStop(LoggerContext context) {

    }

    @Override
    public void onLevelChange(ch.qos.logback.classic.Logger logger, Level level) {

    }

    @Override
    public void start() {
        if (started) {
            return;
        }
        Context context = getContext();
        context.putProperty("maxHistoryDay", getMaxHistoryDay());
        started = true;
    }

    @Override
    public void stop() {

    }

    @Override
    public boolean isStarted() {
        return false;
    }

    private String getMaxHistoryDay() {
        Map<String, String> logMap = null;
        // 如果找不到默认8天
        if (null == logMap) {
            return "8";
        }else {
            return logMap.getOrDefault("log.max.history.days","8");
        }
    }
}

