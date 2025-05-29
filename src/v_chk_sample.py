import sys
import os

def main():
    if len(sys.argv) != 3:
        print("Usage: v_chk.py <vault_dir> <config_dir>")
        sys.exit(1)

    vault_dir = sys.argv[1]
    config_dir = sys.argv[2]

    print(f"Vault Directory: {vault_dir}")
    print(f"Config Directory: {config_dir}")

    config_path = os.path.join(vault_dir, config_dir)
    if os.path.exists(config_path):
        print(f"Contents of '{config_path}':")
        for fname in os.listdir(config_path):
            print(f" - {fname}")
    else:
        print(f"Config directory not found at: {config_path}")

if __name__ == "__main__":
    main()
