from pathlib import (Path)


def main() -> None:
    cfg_list = [
        'E:\\o Misc Vaults\\dummy\\test\\PKM LM',
        'E:/o Misc Vaults/Ideaverse Lite 1.5',
        'E:/o Misc Vaults/freshUlt',
        'E:/o Misc Vaults/DashboardPlusPlus-master',
        'E:/o Misc Vaults/JournalCraft-Obsidian-v100',
        'E:/o Misc Vaults/kepano-obsidian-main',
        'E:/o Misc Vaults/Obs',
        'E:/o Misc Vaults/Obsidian-CSS-Snippets-Collection',
        'E:/o Misc Vaults/obsidian-example-lifeos-main',
        'E:/o Misc Vaults/obsidian-hub-0.2.1',
        'E:/o Misc Vaults/obsidian_dataview_example_vault-main',
        'E:/o Misc Vaults/ObsidianTTRPGVault-main',
        'E:/o Misc Vaults/Personal-Wiki-main',
        'E:/o Misc Vaults/Ultimate Starter Vault 2.2 Beta'
    ]

    for d in cfg_list:
        vault_path = Path(d)
        # if not vault_path.exists():
        #     print(f"Vault path does not exist: {vault_path}")
        #     continue

        print(f"Processing vault: {vault_path}")
        # Here you would call the function to copy the vault to the config
        # sys_cfg.copy_vault_to_config(vault_path)

        # And then run the v_chk system
        # sys_cfg.run_v_chk()
        a = str(vault_path.parent)
        s = str(vault_path)
        print(f"Workbook generated for vault: {vault_path}")

        proceed = input("Press Enter to continue to the next vault, or type 'exit' to quit: ")
        if proceed.lower() == 'exit':
            break


if __name__ == '__main__':
    main()
