import yaml

vault_id = 'v_id'
vault_name = 'v_name'
dir_vault = 'dir_v'

cfg_data = {
              'vault_id':           vault_id
            , 'vault_name':         vault_name
            , 'dir_vault':          dir_vault
}
pn_file = "test.yaml"
with open(pn_file, 'w') as file:
    yaml.dump(cfg_data, file, default_flow_style=False)

with open(pn_file, 'r') as file:
    cfg_data = yaml.safe_load(file)

print( cfg_data.vault_id)