"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const fs = __importStar(require("fs"));
const os = __importStar(require("os"));
const path = __importStar(require("path"));
function getObsidianConfigPath() {
    const platform = os.platform();
    if (platform === 'win32') {
        return path.join(process.env.APPDATA || '', 'Obsidian', 'obsidian.json');
    }
    else if (platform === 'darwin') {
        return path.join(os.homedir(), 'Library', 'Application Support', 'Obsidian', 'obsidian.json');
    }
    else {
        return path.join(os.homedir(), '.config', 'Obsidian', 'obsidian.json');
    }
}
function loadObsidianConfig() {
    const configPath = getObsidianConfigPath();
    if (!fs.existsSync(configPath)) {
        console.error(`Obsidian config not found: ${configPath}`);
        process.exit(1);
    }
    const jsonText = fs.readFileSync(configPath, 'utf8');
    return JSON.parse(jsonText);
}
function listAllVaults(config) {
    if (!config.vaults || Object.keys(config.vaults).length === 0) {
        console.log('No vaults found.');
        return;
    }
    console.log('Available Vaults:\n');
    for (const [vaultId, vaultInfo] of Object.entries(config.vaults)) {
        const configDir = vaultInfo.configDir || '.obsidian';
        const vaultPath = vaultInfo.path;
        console.log(`Vault ID:   ${vaultId}`);
        console.log(`Vault Path: ${vaultPath}`);
        console.log(`Config Dir: ${configDir}\n`);
    }
}
function getVaultConfigDir(config, vaultId) {
    const vaultInfo = config.vaults?.[vaultId];
    if (!vaultInfo) {
        console.error(`Vault ID not found: ${vaultId}`);
        return null;
    }
    return vaultInfo.configDir || '.obsidian';
}
const config = loadObsidianConfig();
const vaultId = process.argv[2];
if (!vaultId) {
    listAllVaults(config);
}
else {
    const result = getVaultConfigDir(config, vaultId);
    if (result !== null) {
        console.log(result);
    }
}
