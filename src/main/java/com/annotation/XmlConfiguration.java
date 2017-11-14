package com.annotation;

import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import com.google.common.collect.Maps;
import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.XMLConfiguration;

import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.TimeUnit;

/**
 * Created by user on 2017/11/14.
 */
public class XmlConfiguration implements Configuration {

    private final static String FIELD_COLUMN_CACHE_PREFIX = "x_";
    private static Cache<String, Map<String, String>> fieldColumnCache = CacheBuilder.newBuilder()
            .expireAfterAccess(1L, TimeUnit.HOURS).build();
    @Override
    public Map<String, String> parse(final Class<?> clazz) throws Exception {
        return fieldColumnCache.get(FIELD_COLUMN_CACHE_PREFIX + clazz.getSimpleName(),
                new Callable<Map<String, String>>() {
                    public Map<String, String> call() throws Exception {
                        Map<String, String> confMap = Maps.newHashMap();
                        XMLConfiguration config = new XMLConfiguration("excel/import/" + clazz.getSimpleName() + ".xml");
                        List<HierarchicalConfiguration> nodes = config.configurationsAt("mapping");
                        for (int i = 0; i < nodes.size(); i++) {
                            HierarchicalConfiguration node = nodes.get(i);
                            String column = node.getString("[@column]");
                            String fieldName = node.getString("[@field]");
                            confMap.put(fieldName, column);
                        }
                        return confMap;
                    }
                });
    }
}
