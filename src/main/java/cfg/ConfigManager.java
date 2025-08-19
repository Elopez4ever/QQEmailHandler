package cfg;

import org.yaml.snakeyaml.Yaml;
import java.io.InputStream;
import java.util.Map;

public class ConfigManager {
    private Map<String, Object> config;

    public ConfigManager() {
        loadConfig();
    }

    private void loadConfig() {
        try {
            Yaml yaml = new Yaml();
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("config.yml");

            if (inputStream == null) {
                System.err.println("找不到配置文件 config.yml，使用默认配置");
                config = Map.of();
                return;
            }

            config = yaml.load(inputStream);
            System.out.println("YAML 配置加载成功!");

        } catch (Exception e) {
            System.err.println("加载 YAML 配置失败: " + e.getMessage());
            config = Map.of();
        }
    }

    public String getString(String path, String defaultValue) {
        Object value = getNestedValue(path);
        return value != null ? value.toString() : defaultValue;
    }

    public int getInt(String path, int defaultValue) {
        Object value = getNestedValue(path);
        if (value instanceof Number) {
            return ((Number) value).intValue();
        }
        return defaultValue;
    }

    public boolean getBoolean(String path, boolean defaultValue) {
        Object value = getNestedValue(path);
        if (value instanceof Boolean) {
            return (Boolean) value;
        }
        return defaultValue;
    }

    @SuppressWarnings("unchecked")
    private Object getNestedValue(String path) {
        String[] keys = path.split("\\.");
        Object current = config;

        for (String key : keys) {
            if (current instanceof Map) {
                current = ((Map<String, Object>) current).get(key);
            } else {
                return null;
            }
        }
        return current;
    }
}
