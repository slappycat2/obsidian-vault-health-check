
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

function getObsidianConfigPath(): string {
  const platform = os.platform();
  if (platform === 'win32') {
    return path.join(process.env.APPDATA || '', 'Obsidian', 'obsidian.json');
  } else if (platform === 'darwin') {
    return path.join(os.homedir(), 'Library', 'Application Support', 'Obsidian', 'obsidian.json');
  } else {
    return path.join(os.homedir(), '.config', 'Obsidian', 'obsidian.json');
  }
}

function loadObsidianConfig(): any {
  const configPath = getObsidianConfigPath();
  if (!fs.existsSync(configPath)) {
    console.error(`Obsidian config not found: ${configPath}`);
    process.exit(1);
  }

  const jsonText = fs.readFileSync(configPath, 'utf8');
  return JSON.parse(jsonText);
}

function listAllVaults(config: any) {
  if (!config.vaults || Object.keys(config.vaults).length === 0) {
    console.log('No vaults found.');
    return;
  }

  console.log('Available Vaults:\n');
  for (const [vaultId, vaultInfo] of Object.entries(config.vaults)) {
    const configDir = (vaultInfo as any).configDir || '.obsidian';
    const vaultPath = (vaultInfo as any).path;
    console.log(`Vault ID:   ${vaultId}`);
    console.log(`Vault Path: ${vaultPath}`);
    console.log(`Config Dir: ${configDir}\n`);
  }
}

function getVaultConfigDir(config: any, vaultId: string): string | null {
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
} else {
  const result = getVaultConfigDir(config, vaultId);
  if (result !== null) {
    console.log(result);
  }
}
