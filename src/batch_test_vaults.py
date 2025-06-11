import os


# 'Ideaverse Lite 1.5 - Copy'
cfg_list = [
              'E:/o Misc Vaults/PKM LM'
            , 'E:/o Misc Vaults/Ideaverse Lite 1.5'
            , 'E:/o Misc Vaults/freshUlt'
            , 'E:/o Misc Vaults/DashboardPlusPlus-master'
            , 'E:/o Misc Vaults/JournalCraft-Obsidian-v100'
            , 'E:/o Misc Vaults/kepano-obsidian-main'
            , 'E:/o Misc Vaults/Obs'
            , 'E:/o Misc Vaults/Obsidian-CSS-Snippets-Collection'
            , 'E:/o Misc Vaults/obsidian-example-lifeos-main'
            , 'E:/o Misc Vaults/obsidian-hub-0.2.1'
            , 'E:/o Misc Vaults/obsidian_dataview_example_vault-main'
            , 'E:/o Misc Vaults/ObsidianTTRPGVault-main'
            , 'E:/o Misc Vaults/Personal-Wiki-main'
            , 'E:/o Misc Vaults/Ultimate Starter Vault 2.2 Beta'
    ]

# This script will use a list of config files and one-by-one, copy them over the the v_chk system CONFIG.yaml file.
# After each copy, it will run the v_chk system to generate a new workbook based on the new config.
# A tkinter widget will pop up after each copy and prompt the user to proceed with the next config file, or exit the script.
def main():
    from src.v_chk_setup import SysConfig
    from src.v_chk_class_lib import Colors

    print(Colors.BOLD + Colors.UNDERLINE + "Batch Test Vaults" + Colors.ENDC)
    print("This script will run the v_chk system on a list of vaults.")
    print("It will copy each vault to the v_chk system CONFIG.yaml file and generate a new workbook.")
    print("Press Enter to continue...")

    input()

    sys_cfg = SysConfig()
    sys_cfg.set_vaults(cfg_list)

    for vault in cfg_list:
        print(f"Processing vault: {vault}")
        sys_cfg.copy_vault_to_config(vault)
        sys_cfg.run_v_chk()
        print(f"Workbook generated for vault: {vault}")

        proceed = input("Press Enter to continue to the next vault, or type 'exit' to quit: ")
        if proceed.lower() == 'exit':
            break