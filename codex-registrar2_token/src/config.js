const path = require('path');
const fs = require('fs');
const { DEFAULT_PHONE_COUNTRIES, normalizePhoneCountries } = require('./phoneCountryCatalog');

const configPath = path.join(__dirname, '..', 'config.json');
const rootDir = path.join(__dirname, '..');

function readJsonConfig(filePath) {
    if (!filePath || !fs.existsSync(filePath)) {
        return {};
    }

    try {
        const content = fs.readFileSync(filePath, 'utf8');
        return JSON.parse(content);
    } catch (error) {
        console.error(`[Config] 解析配置文件失败: ${filePath} -> ${error.message}`);
        return {};
    }
}

function resolveProfileConfigPath() {
    const explicitFile = String(process.env.CONFIG_FILE || '').trim();
    if (explicitFile) {
        return path.isAbsolute(explicitFile)
            ? explicitFile
            : path.resolve(rootDir, explicitFile);
    }

    const profile = String(process.env.CONFIG_PROFILE || '').trim();
    if (profile) {
        return path.join(rootDir, `config.${profile}.json`);
    }

    if (process.platform === 'darwin') {
        return path.join(rootDir, 'config.local.json');
    }
    if (process.platform === 'linux') {
        return path.join(rootDir, 'config.server.json');
    }
    return '';
}

function parseBoolean(value, defaultValue) {
    if (value === undefined || value === null || value === '') return defaultValue;
    if (typeof value === 'boolean') return value;
    const normalized = String(value).trim().toLowerCase();
    if (['1', 'true', 'yes', 'y', 'on'].includes(normalized)) return true;
    if (['0', 'false', 'no', 'n', 'off'].includes(normalized)) return false;
    return defaultValue;
}

function normalizeMailDomains(domains, fallback = '') {
    const list = Array.isArray(domains) ? domains : [];
    const normalized = list
        .map(item => String(item || '').trim().replace(/^@/, ''))
        .filter(Boolean);

    if (normalized.length > 0) return normalized;

    const single = String(fallback || '').trim().replace(/^@/, '');
    return single ? [single] : [];
}

// 读取配置文件：基础 config.json + 环境覆盖文件
function loadConfig() {
    const baseConfig = readJsonConfig(configPath);
    const profileConfigPath = resolveProfileConfigPath();
    const profileConfig = profileConfigPath ? readJsonConfig(profileConfigPath) : {};

    if (profileConfigPath && fs.existsSync(profileConfigPath)) {
        console.log(`[Config] 使用覆盖配置: ${profileConfigPath}`);
    }

    if (!fs.existsSync(configPath) && (!profileConfigPath || !fs.existsSync(profileConfigPath))) {
        console.error(`[Config] 未找到可用配置文件: ${configPath}`);
    }

    return {
        ...baseConfig,
        ...profileConfig,
    };
}

const config = loadConfig();
const phoneCountries = normalizePhoneCountries(
    Array.isArray(config.phoneCountries) && config.phoneCountries.length > 0
        ? config.phoneCountries
        : DEFAULT_PHONE_COUNTRIES
);
const mailDomains = normalizeMailDomains(config.mailDomains, config.mailDomain);

module.exports = {
    // HeroSMS
    heroSmsApiKey: config.heroSmsApiKey,
    heroSmsService: config.heroSmsService || 'dr',
    heroSmsCountry: parseInt(config.heroSmsCountry, 10) || 16,
    heroSmsPromptCountrySelection: parseBoolean(config.heroSmsPromptCountrySelection, true),
    heroSmsCountryTopN: parseInt(config.heroSmsCountryTopN, 10) || 10,

    // Cloudflare 临时邮箱
    mailBaseUrl: config.mailBaseUrl || '',
    mailAdminPassword: config.mailAdminPassword,
    mailSitePassword: config.mailSitePassword || '',
    mailDomain: mailDomains[0] || '',
    mailDomains,
    mailProvider: config.mailProvider || 'cloud-mail', // cloud-mail | legacy | auto
    mailAdminEmail: config.mailAdminEmail || '',
    mailAdminToken: config.mailAdminToken || '',
    mailUserType: parseInt(config.mailUserType, 10) || 1,

    // 代理
    proxyHost: config.proxyHost || '',
    proxyPort: parseInt(config.proxyPort, 10) || 0,
    proxyUsername: config.proxyUsername || '',
    proxyPassword: config.proxyPassword || '',

    // OAuth
    oauthClientId: config.oauthClientId || 'app_EMoamEEZ73f0CkXaXp7hrann',
    oauthRedirectPort: parseInt(config.oauthRedirectPort, 10) || 1455,
    tokenOutputDir: config.tokenOutputDir || '',
    tokenOutputDirs: Array.isArray(config.tokenOutputDirs)
        ? config.tokenOutputDirs.filter(Boolean)
        : [],

    // 浏览器
    useChrome: config.useChrome !== false,
    chromePath: config.chromePath || 'google-chrome-stable',

    // 手机国家
    phoneCountryCode: String(config.phoneCountryCode || 'GB').trim().toUpperCase(),
    phoneCountries,
};
